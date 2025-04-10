const { google } = require('googleapis');
const axios = require('axios');
const cron = require('node-cron');

const BASE_URL = 'https://study.tfshighschool.com/webservice/rest/server.php';
const TOKEN = 'your_moodle_token_here';
const FORMAT = 'json';

async function fetchMoodleData(wsfunction, params) {
    const url = `${BASE_URL}?wstoken=${TOKEN}&wsfunction=${wsfunction}&moodlewsrestformat=${FORMAT}`;
    const response = await axios.get(url, {
        params,
        httpsAgent: new (require('https').Agent)({
            rejectUnauthorized: false, // Temporary workaround for SSL issue
        }),
    });
    return response.data;
}

async function generateAssignmentReport() {
    try {
        console.log('Fetching courses...');
        const coursesData = await fetchMoodleData('core_course_get_courses', {});
        const courses = coursesData.filter(course => course.id !== 1);

        const reportData = [];
        for (const course of courses) {
            console.log(`Fetching assignments for course: ${course.fullname} (${course.id})`);
            const assignmentsData = await fetchMoodleData('mod_assign_get_assignments', { 'courseids[0]': course.id });
            const assignments = assignmentsData.courses[0]?.assignments || [];

            for (const assignment of assignments) {
                console.log(`Fetching submissions for assignment: ${assignment.name} (${assignment.id})`);
                const submissionsData = await fetchMoodleData('mod_assign_get_submissions', { 'assignmentids[0]': assignment.id });
                const submissions = submissionsData.assignments[0]?.submissions || [];

                for (const submission of submissions) {
                    const userData = await fetchMoodleData('core_user_get_users_by_field', { 'field': 'id', 'values[0]': submission.userid });
                    const user = userData[0] || {};

                    reportData.push({
                        CourseName: course.fullname,
                        AssignmentName: assignment.name,
                        StudentName: user.fullname || 'Unknown',
                        SubmissionTime: submission.timemodified ? new Date(submission.timemodified * 1000).toISOString() : 'N/A',
                        Grade: submission.grade || 'Not graded',
                    });
                }
            }
        }

        // Create a Google Sheet directly
        const auth = new google.auth.GoogleAuth({
            scopes: ['https://www.googleapis.com/auth/spreadsheets'], // Only need Sheets API scope
        });
        const sheets = google.sheets({ version: 'v4', auth });

        // Create a new spreadsheet
        const spreadsheet = await sheets.spreadsheets.create({
            requestBody: {
                properties: {
                    title: 'All_Courses_Submissions',
                },
            },
        });

        const spreadsheetId = spreadsheet.data.spreadsheetId;
        console.log(`Created Google Sheet with ID: ${spreadsheetId}`);

        // Prepare data for the sheet
        const headers = ['CourseName', 'AssignmentName', 'StudentName', 'SubmissionTime', 'Grade'];
        const values = [headers, ...reportData.map(row => headers.map(header => row[header] || ''))];

        // Write data to the sheet
        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: 'Sheet1!A1',
            valueInputOption: 'RAW',
            requestBody: {
                values,
            },
        });

        // Make the sheet publicly readable
        const drive = google.drive({ version: 'v3', auth });
        await drive.permissions.create({
            fileId: spreadsheetId,
            requestBody: {
                role: 'reader',
                type: 'anyone',
            },
        });

        const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
        console.log(`Report generated and uploaded to Google Sheets: ${sheetUrl}`);
        return sheetUrl;
    } catch (error) {
        console.error('Error generating report:', error);
        throw error;
    }
}

// Express route to trigger report generation
app.get('/generate-report', async (req, res) => {
    try {
        const sheetUrl = await generateAssignmentReport();
        res.status(200).json({ message: 'Report generated successfully', url: sheetUrl });
    } catch (error) {
        res.status(500).json({ message: 'Error generating report', error: error.message });
    }
});

// Schedule the report generation to run every day at midnight (America/New_York)
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
    timezone: 'America/New_York'
});

app.listen(4005, () => {
    console.log('Server running at http://localhost:4005');
    console.log('Cron job scheduled to run generateAssignmentReport every day at midnight');
});