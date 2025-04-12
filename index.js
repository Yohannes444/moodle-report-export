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
const FORMAT = config.format;
const DAYS_BACK = config.daysBack || 600; // Default to 600 days if not specified

// Helper function to make API calls
async function fetchMoodleData(baseUrl, token, wsfunction, params) {
    try {
        const url = `${baseUrl}?wstoken=${token}&wsfunction=${wsfunction}&moodlewsrestformat=${FORMAT}`;
        const response = await axios.get(url, {
            params,
            httpsAgent: new https.Agent({
                rejectUnauthorized: false, // Disable SSL verification
            }),
        });

        if (!response.data) {
            throw new Error(`Empty response from Moodle API for ${wsfunction}`);
        }

        if (response.data.errorcode || response.data.exception) {
            throw new Error(`Moodle API error: ${response.data.message || response.data.exception}`);
        }

        return response.data;
    } catch (error) {
        console.error(`Error fetching Moodle data for ${wsfunction}:`, error.message);
        if (error.response) {
            console.error('Response data:', error.response.data);
        }
        throw error;
    }
}

// Google Sheets API setup for a specific service account
function getGoogleApis(serviceAccountKey) {
    const auth = new google.auth.GoogleAuth({
        keyFile: serviceAccountKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'],
    });
    return {
        sheets: google.sheets({ version: 'v4', auth }),
        drive: google.drive({ version: 'v3', auth })
    };
}

// Function to upload .xlsx file to Google Drive and convert to Google Sheets
// Function to upload .xlsx file to Google Drive and convert to Google Sheets
async function uploadToGoogleSheets(drive, filePath, schoolName) {
    try {
        // Format the date as YYYY-MM-DD for cleaner file names
        const formattedDate = new Date().toISOString().split('T')[0];
        const fileMetadata = {
            name: `Ungraded_Submissions_${schoolName}_${formattedDate}`,
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

        console.log(`File uploaded to Google Sheets for ${schoolName}:`, response.data.webViewLink);
        return response.data.webViewLink;
    } catch (error) {
        console.error(`Error uploading to Google Sheets for ${schoolName}:`, error.message);
        throw error;
    }
}

// Function to generate report for ungraded assignments and quizzes for a single school
async function generateAssignmentReportForSchool(school) {
    const { name, baseUrl, token, serviceAccountKey } = school;
    console.log(`Generating report for school: ${name}`);
    
    try {
        const wb = new xl.Workbook();
        const ws = wb.addWorksheet('Ungraded Submissions');

        // Set column widths
        ws.column(1).setWidth(10);  // Course ID
        ws.column(2).setWidth(30);  // Submission Name
        ws.column(3).setWidth(20);  // Student Name
        ws.column(4).setWidth(20);  // Student Username
        ws.column(5).setWidth(30);  // Student Email
        ws.column(6).setWidth(25);  // Date Submitted
        ws.column(7).setWidth(75);  // Direct Link to Submission

        // Set headers
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

        console.log(`Fetching courses for ${name}...`);
        const coursesResponse = await fetchMoodleData(baseUrl, token, 'core_course_get_courses', {});
        const courses = Array.isArray(coursesResponse) ? coursesResponse : [];
        console.log(`Found ${courses.length} courses for ${name}`);

        const coursePromises = courses
            .filter(course => course && course.id !== 1)
            .map(course => 
                fetchMoodleData(baseUrl, token, 'core_course_get_contents', { courseid: course.id })
                    .then(contents => ({ courseId: course.id, shortname: course.shortname, contents }))
                    .catch(err => {
                        console.error(`Error fetching contents for course ${course.id} in ${name}:`, err.message);
                        return { courseId: course.id, contents: [] };
                    })
            );
        const courseResults = await Promise.all(coursePromises);
        console.log(`Fetched contents for all courses in ${name}`);

        // Calculate time threshold for filtering submissions (X days ago)
        const timeThreshold = Math.floor(Date.now() / 1000) - (DAYS_BACK * 24 * 60 * 60);

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
                    fetchMoodleData(baseUrl, token, 'mod_assign_get_submissions', { 'assignmentids': assignmentIds, "status": "submitted" })
                        .then(submissionsData => ({ type: 'assign', items: assignments, data: submissionsData }))
                        .catch(err => {
                            console.error(`Error fetching submissions for course ${courseId} in ${name}:`, err.message);
                            return { type: 'assign', items: assignments, data: { assignments: [] } };
                        })
                );
            }

            if (quizzes.length > 0) {
                const quizIds = quizzes.map(q => q.id);
                modulePromises.push(
                    fetchMoodleData(baseUrl, token, 'mod_quiz_get_user_attempts', { 'quizids': quizIds, 'status': 'finished' })
                        .then(async attemptsData => {
                            const ungradedAttempts = [];
                            for (const attempt of attemptsData.attempts || []) {
                                const gradeData = await fetchMoodleData(baseUrl, token, 'mod_quiz_get_user_best_grade', {
                                    quizid: attempt.quiz,
                                    userid: attempt.userid
                                }).catch(err => {
                                    console.error(`Error fetching grade for quiz ${attempt.quiz}, user ${attempt.userid} in ${name}:`, err.message);
                                    return { hasgrade: false };
                                });
                                if (!gradeData.hasgrade) {
                                    ungradedAttempts.push(attempt);
                                }
                            }
                            return { type: 'quiz', items: quizzes, data: { attempts: ungradedAttempts } };
                        }).catch(err => {
                            console.error(`Error fetching quiz attempts for course ${courseId} in ${name}:`, err.message);
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
                        console.log(`Checking assignment: ${assignment.name}, User: ${submission.userid}, GradingStatus: ${submission.gradingstatus}, Time: ${new Date(submission.timemodified * 1000).toISOString()}`);
                        if ((submission.gradingstatus === 'notgraded') && 
                            submission.timemodified >= timeThreshold) {
                            console.log(`Including assignment submission: ${assignment.name}, User: ${submission.userid}, Submitted: ${new Date(submission.timemodified * 1000).toISOString()}`);
                            allSubmissions.push({
                                courseId: assignment.shortname,
                                submissionName: assignment.name,
                                cmid: assignment.cmid,
                                studentId: submission.userid,
                                dateSubmitted: new Date(submission.timemodified * 1000).toISOString(),
                                type: 'assign'
                            });
                        }
                    });
                });
            } else if (type === 'quiz' && data.attempts) {
                data.attempts.forEach(attempt => {
                    const quiz = items.find(q => q.id === attempt.quiz);
                    if (attempt.timefinish >= timeThreshold) {
                        console.log(`Including quiz attempt: ${quiz.name}, User: ${attempt.userid}, Submitted: ${new Date(attempt.timefinish * 1000).toISOString()}`);
                        allSubmissions.push({
                            courseId: quiz.shortname,
                            submissionName: quiz.name,
                            cmid: quiz.cmid,
                            studentId: attempt.userid,
                            dateSubmitted: new Date(attempt.timefinish * 1000).toISOString(),
                            type: 'quiz'
                        });
                    }
                });
            }
        }

        if (allSubmissions.length === 0) {
            console.log(`No ungraded submissions found within the last ${DAYS_BACK} days for ${name}.`);
            ws.cell(2, 1).string(`No ungraded submissions found within the last ${DAYS_BACK} days for ${name}.`);
        } else {
            const uniqueStudentIds = [...new Set(allSubmissions.map(s => s.studentId))];
            const userPromises = uniqueStudentIds.map(studentId =>
                fetchMoodleData(baseUrl, token, 'core_user_get_users_by_field', {
                    field: 'id',
                    'values[0]': studentId
                }).then(users => ({ studentId, user: users[0] || {} }))
                .catch(err => {
                    console.error(`Error fetching user ${studentId} in ${name}:`, err.message);
                    return { studentId, user: {} };
                })
            );
            const userResults = await Promise.all(userPromises);
            const userMap = new Map(userResults.map(r => [r.studentId, r.user]));

            for (const submission of allSubmissions) {
                console.log(`Processing ungraded submission for student ${submission.studentId} in ${name}`);
                const student = userMap.get(submission.studentId) || {};
                ws.cell(row, 1).string(submission.courseId);
                ws.cell(row, 2).string(submission.submissionName);
                ws.cell(row, 3).string(student.fullname || 'Unknown');
                ws.cell(row, 4).string(student.username || 'Unknown');
                ws.cell(row, 5).string(student.email || 'Unknown');
                ws.cell(row, 6).string(submission.dateSubmitted);
                ws.cell(row, 7).string(
                    `${baseUrl.replace('/webservice/rest/server.php', '')}/mod/${submission.type}/view.php?id=${submission.cmid}&rownum=0&action=grader&userid=${submission.studentId}`
                );
                row++;
            }
        }

        // Use school name for local file to avoid conflicts
        const formattedDate = new Date().toISOString().split('T')[0];
        const filePath = `Ungraded_Submissions_${name}_${formattedDate}.xlsx`;
        await new Promise((resolve, reject) => {
            wb.write(filePath, (err) => {
                if (err) reject(err);
                else resolve();
            });
        });
        console.log(`Excel file generated for ${name}:`, filePath);

        const { drive } = getGoogleApis(serviceAccountKey);
        const googleSheetLink = await uploadToGoogleSheets(drive, filePath, name);
        console.log(`Report generated and uploaded to Google Sheets for ${name}: ${googleSheetLink}`);

        fs.unlink(filePath, (err) => {
            if (err) {
                console.error(`Error deleting local Excel file for ${name}:`, err.message);
            } else {
                console.log(`Local Excel file deleted successfully for ${name}:`, filePath);
            }
        });

        return googleSheetLink;
    } catch (error) {
        console.error(`Error generating report for ${name}:`, error.message);
        if (error.response) {
            console.error('Response data:', error.response.data);
        }
        throw error;
    }
}

// Function to generate reports for all schools
async function generateAssignmentReport() {
    const results = [];
    for (const school of config.schools) {
        try {
            const link = await generateAssignmentReportForSchool(school);
            results.push(`Report for ${school.name}: ${link}`);
        } catch (error) {
            results.push(`Failed to generate report for ${school.name}: ${error.message}`);
        }
    }
    return results.join('\n');
}

// Express route to trigger report generation manually
app.get('/generate-report', async (req, res) => {
    try {
        const message = await generateAssignmentReport();
        res.send(`Reports generated:\n${message}`);
    } catch (error) {
        res.status(500).send('Error generating reports');
    }
});

app.get("/", async (req, res) => {
    res.send(`hello`);
});

// Schedule the report to run every day at midnight (00:00)
cron.schedule('0 0 * * *', async () => {
    console.log('Running scheduled report generation at', new Date().toISOString());
    try {
        const message = await generateAssignmentReport();
        console.log('Scheduled report generation completed:\n', message);
    } catch (error) {
        console.error('Error in scheduled report generation:', error);
    }
}, {
    scheduled: true,
    timezone: 'America/New_York'
});

// Start server
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
    console.log(`Cron job scheduled to run generateAssignmentReport every day at midnight`);
});