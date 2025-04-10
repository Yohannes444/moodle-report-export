const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const https = require("https");
const cron = require('node-cron');

const app = express();
const port = 4005;

const BASE_URL = 'https://study.tfshighschool.com/webservice/rest/server.php';
const TOKEN = '2930400cf30439f2aed3774c924c3669';
const FORMAT = 'json';

// Helper function to make API calls
async function fetchMoodleData(wsfunction, params) {
    const url = `${BASE_URL}?wstoken=${TOKEN}&wsfunction=${wsfunction}&moodlewsrestformat=${FORMAT}`;
    try {
        const response = await axios.get(url, {
            params,
            httpsAgent: new https.Agent({
                rejectUnauthorized: false, // Disable SSL verification
            }),
        });
        return response.data;
    } catch (error) {
        console.error(`Error fetching Moodle data for ${wsfunction}:`, error.response ? error.response.data : error.message);
        throw error;
    }
}

// Google Sheets API setup with ADC
const auth = new google.auth.GoogleAuth({
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    projectId: '928317390511', // Explicitly set the project ID to avoid mismatches
});

const sheets = google.sheets({ version: 'v4', auth });

// Function to generate report for all courses (assignments and quizzes)
async function generateAssignmentReport() {
    try {
        console.log('Fetching courses...');
        const courses = await fetchMoodleData('core_course_get_courses', {});
        console.log(`Found ${courses.length} courses`);

        const coursePromises = courses
            .filter(course => course.id !== 1)
            .map(course => 
                fetchMoodleData('core_course_get_contents', { courseid: course.id })
                    .then(contents => ({ courseId: course.id, shortname: course.shortname, contents }))
                    .catch(err => {
                        console.error(`Error fetching contents for course ${course.id}:`, err.message);
                        return { courseId: course.id, contents: [] };
                    })
            );
        const courseResults = await Promise.all(coursePromises);
        console.log('Fetched contents for all courses');

        const modulePromises = [];
        for (const { courseId, shortname, contents } of courseResults) {
            const assignments = [];
            const quizzes = [];
            contents.forEach(section => {
                section.modules.forEach(module => {
                    if (module.modname === 'assign') {
                        assignments.push({
                            id: module.instance,
                            name: module.name,
                            cmid: module.id,
                            courseId,
                            shortname
                        });
                    } else if (module.modname === 'quiz') {
                        quizzes.push({
                            id: module.instance,
                            name: module.name,
                            cmid: module.id,
                            courseId,
                            shortname
                        });
                    }
                });
            });

            if (assignments.length > 0) {
                const assignmentIds = assignments.map(a => a.id);
                modulePromises.push(
                    fetchMoodleData('mod_assign_get_submissions', { 'assignmentids': assignmentIds })
                        .then(submissionsData => ({ type: 'assign', items: assignments, data: submissionsData }))
                        .catch(err => {
                            console.error(`Error fetching submissions for course ${courseId}:`, err.message);
                            return { type: 'assign', items: assignments, data: { assignments: [] } };
                        })
                );
            }

            if (quizzes.length > 0) {
                const quizIds = quizzes.map(q => q.id);
                modulePromises.push(
                    fetchMoodleData('mod_quiz_get_user_attempts', { 'quizids': quizIds, 'status': 'finished' })
                        .then(attemptsData => ({ type: 'quiz', items: quizzes, data: attemptsData }))
                        .catch(err => {
                            console.error(`Error fetching quiz attempts for course ${courseId}:`, err.message);
                            return { type: 'quiz', items: quizzes, data: { attempts: [] } };
                        })
                );
            }
        }
        const moduleResults = await Promise.all(modulePromises);

        const allSubmissions = [];
        for (const { type, items, data } of moduleResults) {
            if (type === 'assign' && data.assignments) {
                data.assignments.forEach(assignmentData => {
                    const assignment = items.find(a => a.id === assignmentData.assignmentid);
                    const submissionList = assignmentData.submissions || [];
                    submissionList.forEach(submission => {
                        allSubmissions.push({
                            courseId: assignment.shortname,
                            submissionName: assignment.name,
                            cmid: assignment.cmid,
                            studentId: submission.userid,
                            dateSubmitted: new Date(submission.timemodified * 1000).toISOString(),
                            type: 'assign'
                        });
                    });
                });
            } else if (type === 'quiz' && data.attempts) {
                data.attempts.forEach(attempt => {
                    const quiz = items.find(q => q.id === attempt.quiz);
                    allSubmissions.push({
                        courseId: quiz.shortname,
                        submissionName: quiz.name,
                        cmid: quiz.cmid,
                        studentId: attempt.userid,
                        dateSubmitted: new Date(attempt.timefinish * 1000).toISOString(),
                        type: 'quiz'
                    });
                });
            }
        }

        const uniqueStudentIds = [...new Set(allSubmissions.map(s => s.studentId))];
        const userPromises = uniqueStudentIds.map(studentId =>
            fetchMoodleData('core_user_get_users_by_field', {
                field: 'id',
                'values[0]': studentId
            }).then(users => ({ studentId, user: users[0] || {} }))
            .catch(err => {
                console.error(`Error fetching user ${studentId}:`, err.message);
                return { studentId, user: {} };
            })
        );
        const userResults = await Promise.all(userPromises);
        const userMap = new Map(userResults.map(r => [r.studentId, r.user]));

        // Prepare data for Google Sheets
        const headers = [
            'Course ID',
            'Submission Name',
            'Student Name',
            'Student Username',
            'Student Email',
            'Date Submitted',
            'Direct Link to Submission'
        ];
        const values = [headers]; // Start with headers

        for (const submission of allSubmissions) {
            console.log(`Processing submission for student ${submission.studentId}`);
            const student = userMap.get(submission.studentId) || {};
            const row = [
                submission.courseId,
                submission.submissionName,
                student.fullname || 'Unknown',
                student.username || 'Unknown',
                student.email || 'Unknown',
                submission.dateSubmitted,
                `${BASE_URL.replace('/webservice/rest/server.php', '')}/mod/${submission.type}/view.php?id=${submission.cmid}`
            ];
            values.push(row);
        }

        // Log authentication details for debugging
        const authClient = await auth.getClient();
        console.log('Auth client type:', authClient.constructor.name);
        console.log('Requested scopes:', authClient.scopes);
        console.log('Project ID:', authClient.projectId || 'Not set');

        // Create a new spreadsheet
        const spreadsheet = await sheets.spreadsheets.create({
            requestBody: {
                properties: {
                    title: 'All_Courses_Submissions_' + new Date().toISOString(),
                },
            },
        });

        const spreadsheetId = spreadsheet.data.spreadsheetId;
        console.log(`Created Google Sheet with ID: ${spreadsheetId}`);

        // Write data to the sheet
        if (values.length === 1) { // Only headers, no data
            console.log('No submissions found for any assignments or quizzes across all courses.');
            values.push(['No submissions found']);
        }

        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: 'Sheet1!A1',
            valueInputOption: 'RAW',
            requestBody: {
                values,
            },
        });

        const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
        console.log(`Report generated and uploaded to Google Sheets: ${sheetUrl}`);
        return sheetUrl;
    } catch (error) {
        console.error('Error generating report:', error.response ? error.response.data : error.message);
        if (error.response && error.response.data && error.response.data.error) {
            console.error('Error details:', error.response.data.error.details || error.response.data.error);
            console.error('Error details:', error.response.data.error.errors || error.response.data.error);
        }
        throw error;
    }
}

// Express route to trigger report generation manually
app.get('/generate-report', async (req, res) => {
    try {
        const message = await generateAssignmentReport();
        res.send(`Report generated and uploaded to Google Sheets: ${message}`);
    } catch (error) {
        res.status(500).send('Error generating report');
    }
});

app.get("/", async (req, res) => {
    res.send(`hello`);
});

// Schedule the report to run every day at midnight (00:00)
cron.schedule('0 0 * * *', async () => {
    console.log('Running scheduled report generation at', new Date().toISOString());
    try {
        await generateAssignmentReport();
        console.log('Scheduled report generation completed successfully');
    } catch (error) {
        console.error('Error in scheduled report generation:', error);
    }
}, {
    scheduled: true,
    timezone: 'America/New_York' // Adjust to your timezone
});

// Start server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
    console.log('Cron job scheduled to run generateAssignmentReport every day at midnight');
});