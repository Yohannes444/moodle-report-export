const express = require('express');
const axios = require('axios');
const xl = require('excel4node');
const { google } = require('googleapis');
const cors = require("cors");
const fs = require('fs');
const https = require("https");
const config = require('./config.json');
const cron = require('node-cron');
const mongoose = require('mongoose');
const Report = require('./models/Report');
const School = require('./models/School');
const path = require("path");
const bodyParser = require("body-parser");
const { getStatistics } = require('./service/staticservice');
const app = express();
app.use(cors({
  origin: ['*','https://moodle-student-hub.lovable.app/', "https://id-preview--742b57b4-885f-4866-8211-d42906db0762.lovable.app",'http://localhost:8080/','https://preview--moodle-student-hub.lovable.app/'],
  methods: ['GET', 'POST', 'PUT', 'DELETE'],
  allowedHeaders: ['Content-Type', 'Authorization'],
}));
const port = config.port;
const FORMAT = config.format;
const DAYS_BACK = config.daysBack || 600;

// Connect to MongoDB
mongoose.connect('mongodb+srv://yohannesmulat777:pu7nRPz0rTeXGFuF@yohannes444.e23yh9p.mongodb.net/moodleReports', {
  useNewUrlParser: true,
  useUnifiedTopology: true
}).then(() => {
  console.log('Connected to MongoDB');
}).catch(err => {
  console.error('MongoDB connection error:', err.message);
});

// Helper function to make API calls
async function fetchMoodleData(baseUrl, token, wsfunction, params) {
    try {
        const url = `${baseUrl}?wstoken=${token}&wsfunction=${wsfunction}&moodlewsrestformat=${FORMAT}`;
        const response = await axios.get(url, {
            params,
            httpsAgent: new https.Agent({
                rejectUnauthorized: false,
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

// Function to upload .lsx file to Google Drive and convert to Google Sheets
async function uploadToGoogleSheets(drive, filePath, schoolName, existingFileId = null) {
    try {
        const fileMetadata = {
            name: `Ungraded_Submissions_${schoolName}`,
            mimeType: 'application/vnd.google-apps.spreadsheet',
        };
        
        const media = {
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            body: fs.createReadStream(filePath),
        };

        let response;
        if (existingFileId) {
            response = await drive.files.update({
                fileId: existingFileId,
                resource: fileMetadata,
                media: media,
                fields: 'id, webViewLink',
            });
        } else {
            response = await drive.files.create({
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
        }

        console.log(`File ${existingFileId ? 'updated' : 'uploaded'} to Google Sheets for ${schoolName}:`, response.data.webViewLink);
        return { id: response.data.id, link: response.data.webViewLink };
    } catch (error) {
        console.error(`Error ${existingFileId ? 'updating' : 'uploading'} to Google Sheets for ${schoolName}:`, error.message);
        throw error;
    }
}

// Function to find existing Google Sheet by name
async function findExistingSheet(drive, schoolName) {
    try {
        const response = await drive.files.list({
            q: `name='Ungraded_Submissions_${schoolName}' and mimeType='application/vnd.google-apps.spreadsheet'`,
            fields: 'files(id, webViewLink)',
        });

        return response.data.files.length > 0 ? response.data.files[0] : null;
    } catch (error) {
        console.error(`Error finding existing sheet for ${schoolName}:`, error.message);
        return null;
    }
}

// Function to generate report for ungraded assignments and quizzes for a single school
async function generateAssignmentReportForSchool(school) {
    const { name, baseUrl, token, serviceAccountKey } = school;
    console.log(`Generating report for school: ${name}`);
    
    const wb = new xl.Workbook();
    let ws;
    
    try {
        ws = wb.addWorksheet('Ungraded Submissions');

        ws.column(1).setWidth(10);
        ws.column(2).setWidth(30);
        ws.column(3).setWidth(20);
        ws.column(4).setWidth(20);
        ws.column(5).setWidth(30);
        ws.column(6).setWidth(25);
        ws.column(7).setWidth(75);

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
                        if ((submission.gradingstatus === 'notgraded') && 
                            submission.timemodified >= timeThreshold) {
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

        let submissions = [];
        let errorMessage = null;
        if (allSubmissions.length === 0) {
            console.log(`No ungraded submissions found within the last ${DAYS_BACK} days for ${name}.`);
            ws.cell(2, 1).string(`No ungraded submissions found within the last ${DAYS_BACK} days for ${name}.`);
            errorMessage = `No ungraded submissions found within the last ${DAYS_BACK} days for ${name}.`;
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
                const student = userMap.get(submission.studentId) || {};
                const submissionData = {
                    courseId: submission.courseId,
                    submissionName: submission.submissionName,
                    studentName: student.fullname || 'Unknown',
                    studentUsername: student.username || 'Unknown',
                    studentEmail: student.email || 'Unknown',
                    dateSubmitted: new Date(submission.dateSubmitted),
                    directLink: `${baseUrl.replace('/webservice/rest/server.php', '')}/mod/${submission.type}/view.php?id=${submission.cmid}&rownum=0&action=grader&userid=${submission.studentId}`
                };
                ws.cell(row, 1).string(submissionData.courseId);
                ws.cell(row, 2).string(submissionData.submissionName);
                ws.cell(row, 3).string(submissionData.studentName);
                ws.cell(row, 4).string(submissionData.studentUsername);
                ws.cell(row, 5).string(submissionData.studentEmail);
                ws.cell(row, 6).string(submissionData.dateSubmitted.toISOString());
                ws.cell(row, 7).string(submissionData.directLink);
                submissions.push(submissionData);
                row++;
            }
        }

        const filePath = `Ungraded_Submissions_${name}.xlsx`;
        await new Promise((resolve, reject) => {
            wb.write(filePath, (err) => {
                if (err) reject(err);
                else resolve();
            });
        });
        console.log(`Excel file generated for ${name}:`, filePath);

        const { drive } = getGoogleApis(serviceAccountKey);
        const existingSheet = await findExistingSheet(drive, name);
        const { id, link } = await uploadToGoogleSheets(drive, filePath, name, existingSheet ? existingSheet.id : null);

        // Check for existing report in MongoDB
        let report = await Report.findOne({ schoolName: name });
        if (report) {
            // Update existing report
            report.submissions = submissions;
            report.errorMessage = errorMessage;
            report.googleSheetsLink = link;
            report.fileId = id;
            await report.save();
            console.log(`Updated MongoDB report for ${name}`);
        } else {
            // Create new report
            report = new Report({
                schoolName: name,
                submissions,
                errorMessage,
                googleSheetsLink: link,
                fileId: id
            });
            await report.save();
            console.log(`Created new MongoDB report for ${name}`);
        }

        fs.unlink(filePath, (err) => {
            if (err) {
                console.error(`Error deleting local Excel file for ${name}:`, err.message);
            } else {
                console.log(`Local Excel file deleted successfully for ${name}:`, filePath);
            }
        });

        console.log(`Report ${existingSheet ? 'updated' : 'generated and uploaded'} to Google Sheets for ${name}: ${link}`);
        return link;
    } catch (error) {
        console.error(`Error generating report for ${name}:`, error.message);
        if (error.response) {
            console.error('Response data:', error.response.data);
        }
        if (ws) {
            ws.cell(2, 1).string(`Error generating report: ${error.message}`);
            const filePath = `Ungraded_Submissions_${name}_Error.xlsx`;
            await new Promise((resolve, reject) => {
                wb.write(filePath, (err) => {
                    if (err) reject(err);
                    else resolve();
                });
            });
            console.log(`Error report generated for ${name}:`, filePath);
            
            const { drive } = getGoogleApis(serviceAccountKey);
            const existingSheet = await findExistingSheet(drive, name);
            const { id, link } = await uploadToGoogleSheets(drive, filePath, name, existingSheet ? existingSheet.id : null);
            
            // Save error report to MongoDB
            let report = await Report.findOne({ schoolName: name });
            const errorMessage = `Error generating report: ${error.message}`;
            if (report) {
                report.submissions = [];
                report.errorMessage = errorMessage;
                report.googleSheetsLink = link;
                report.fileId = id;
                await report.save();
                console.log(`Updated MongoDB error report for ${name}`);
            } else {
                report = new Report({
                    schoolName: name,
                    submissions: [],
                    errorMessage,
                    googleSheetsLink: link,
                    fileId: id
                });
                await report.save();
                console.log(`Created new MongoDB error report for ${name}`);
            }

            fs.unlink(filePath, (err) => {
                if (err) {
                    console.error(`Error deleting local error Excel file for ${name}:`, err.message);
                }
            });
            
            return link;
        }
        throw error;
    }
}

// Function to generate reports for all schools
async function generateAssignmentReport() {
    const results = [];
    const schools = await School.find();
    for (const school of schools) {
        try {
            const link = await generateAssignmentReportForSchool(school);
            results.push(`Report for ${school.name}: ${link}`);
        } catch (error) {
            results.push(`Failed to generate report for ${school.name}: ${error.message}`);
        }
    }
    return results.join('\n');
}

app.use(express.static(path.join(__dirname, 'public')));
app.use(bodyParser.urlencoded({ extended: true, limit: "50mb" }));
app.use(bodyParser.json({ limit: "50mb" }));

// Express route to trigger report generation manually
app.get('/generate-report', async (req, res) => {
    try {
        const message = await generateAssignmentReport();
        res.send(`Reports generated:\n${message}`);
    } catch (error) {
        res.status(500).send('Error generating reports');
    }
});

app.get("/statistics",getStatistics)

app.post("/create-school", async (req, res) => {
    try {
        console.log("Creating school:", req.body);
        //payload validation
        if (!req.body.name || !req.body.baseUrl || !req.body.token ) {
            return res.status(400).json({ message: "Missing required fields" });
        }
        const { name, baseUrl, token, serviceAccountKey } = req.body;
        const school = new School({ name, baseUrl, token, serviceAccountKey });
        await school.save();
        res.status(200).json(school);
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

// Endpoint to fetch all reports
app.get('/reports', async (req, res) => {
    try {
        console.log('Fetching all reports');
        const reports = await Report.find({}).lean();
        const sortedReports = reports.map(report => ({
            schoolName: report.schoolName,
            submissions: report.submissions.sort((a, b) => b.dateSubmitted - a.dateSubmitted),
            errorMessage: report.errorMessage,
            googleSheetsLink: report.googleSheetsLink,
            createdAt: report.createdAt,
            updatedAt: report.updatedAt
        }));
        res.json(sortedReports);
    } catch (error) {
        console.error('Error fetching all reports:', error.message);
        res.status(500).json({ message: 'Error fetching reports' });
    }
});
// Endpoint to fetch report for a specific school
app.get('/reports/:schoolName', async (req, res) => {
    try {
      const { schoolName } = req.params;
      const { submissionName } = req.query; // Get filter from query parameter
      let query = { schoolName };
  
      const report = await Report.findOne(query);
      if (!report) {
        return res.status(404).json({ message: `No report found for school: ${schoolName}` });
      }
  
      // Apply filter if submissionName is provided
      let filteredSubmissions = report.submissions;
      if (submissionName) {
        filteredSubmissions = report.submissions.filter((submission) =>
          submission.submissionName.toLowerCase().includes(submissionName.toLowerCase())
        );
      }
  
      res.json({
        schoolName: report.schoolName,
        submissions: filteredSubmissions,
        errorMessage: report.errorMessage,
        googleSheetsLink: report.googleSheetsLink,
        createdAt: report.createdAt,
        updatedAt: report.updatedAt
      });
    } catch (error) {
      console.error(`Error fetching report for ${req.params.schoolName}:`, error.message);
      res.status(500).json({ message: 'Error fetching report' });
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