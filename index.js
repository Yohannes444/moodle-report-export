const express = require('express');
const axios = require('axios');
const xl = require('excel4node');
const { google } = require('googleapis');
const fs = require('fs');
const https = require("https");
const config = require('./config.json');
const cron = require('node-cron');

const app = express();
const port = config.port;
const BASE_URL = config.baseUrl;
const TOKEN = config.token;
const FORMAT = config.format;


// Helper function to make API calls
async function fetchMoodleData(wsfunction, params) {
    const url = `${BASE_URL}?wstoken=${TOKEN}&wsfunction=${wsfunction}&moodlewsrestformat=${FORMAT}`;
    const response = await axios.get(url, {
        params,
        httpsAgent: new https.Agent({
            rejectUnauthorized: false, // Disable SSL verification
          }),
    });
    return response.data;
}

// Google Sheets API setup
const auth = new google.auth.GoogleAuth({
    keyFile: config.serviceAccountKey, // Use config instead of process.env
    scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'],
});

const sheets = google.sheets({ version: 'v4', auth });
const drive = google.drive({ version: 'v3', auth });

// Function to upload .xlsx file to Google Drive and convert to Google Sheets
async function uploadToGoogleSheets(filePath) {
    try {
        const fileMetadata = {
            name: 'All_Courses_Submissions_' + new Date().toISOString() + '.xlsx',
            mimeType: 'application/vnd.google-apps.spreadsheet',
        };
        const media = {
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            body: fs.createReadStream(filePath),
        };

        const response = await drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id, webViewLink',
        });

        await drive.permissions.create({
            fileId: response.data.id,
            requestBody: {
                role: 'reader',
                type: 'anyone',
            },
        });

        console.log('File uploaded to Google Sheets and made public:', response.data.webViewLink);
        return response.data.webViewLink;
    } catch (error) {
        console.error('Error uploading to Google Sheets:', error.message);
        throw error;
    }
}

// Function to generate report for all courses (assignments and quizzes)
// ... (previous code remains the same until generateAssignmentReport function)

async function generateAssignmentReport() {
    try {
        const wb = new xl.Workbook();
        const ws = wb.addWorksheet('Submissions');

        ws.column(1).setWidth(10);  // Course ID
        ws.column(2).setWidth(30);  // Submission Name
        ws.column(3).setWidth(20);  // Student Name
        ws.column(4).setWidth(20);  // Student Username
        ws.column(5).setWidth(30);  // Student Email
        ws.column(6).setWidth(25);  // Date Submitted
        ws.column(7).setWidth(50);  // Direct Link to Submission

        const headers = [
            'Course ID',
            'Submission Name',
            'Student Name',
            'Student Username',
            'Student Email',
            'Date Submitted',
            'Direct Link to Submission'
        ];
        headers.forEach((header, index) => ws.cell(1, index + 1).string(header));
        let row = 2;
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

        for (const submission of allSubmissions) {
            console.log(`Processing submission for student ${submission.studentId}`);
            const student = userMap.get(submission.studentId) || {};
            ws.cell(row, 1).string(submission.courseId);
            ws.cell(row, 2).string(submission.submissionName);
            ws.cell(row, 3).string(student.fullname || 'Unknown');
            ws.cell(row, 4).string(student.username || 'Unknown');
            ws.cell(row, 5).string(student.email || 'Unknown');
            ws.cell(row, 6).string(submission.dateSubmitted);
            ws.cell(row, 7).string(
                `${BASE_URL.replace('/webservice/rest/server.php', '')}/mod/${submission.type}/view.php?id=${submission.cmid}`
            );
            row++;
        }

        const filePath = 'All_Courses_Submissions.xlsx';
        if (row === 2) {
            console.log('No submissions found for any assignments or quizzes across all courses.');
            ws.cell(2, 1).string('No submissions found');
        }
        await new Promise((resolve, reject) => {
            wb.write(filePath, (err) => {
                if (err) reject(err);
                else resolve();
            });
        });
        console.log('Excel file generated:', filePath);

        const googleSheetLink = await uploadToGoogleSheets(filePath);
        console.log(`Report generated and uploaded to Google Sheets: ${googleSheetLink}`);

        // Delete the local Excel file after successful upload
        fs.unlink(filePath, (err) => {
            if (err) {
                console.error('Error deleting local Excel file:', err.message);
            } else {
                console.log('Local Excel file deleted successfully:', filePath);
            }
        });

        return googleSheetLink;
    } catch (error) {
        console.error('Error generating report:', error.response ? error.response.data : error.message);
        console.error('Error details:', error.response ? error.response.data.error.details : error.message);
        console.error('Error details:', error.response ? error.response.data.error.errors : error.message);
        throw error;
    }
}

// ... (rest of the code remains the same)

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
        res.send(`hellow `);
 
    })

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