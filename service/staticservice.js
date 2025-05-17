const mongoose = require('mongoose');
const Report = require('../models/Report');
const School = require('../models/School');

const getStatistics = async (req, res) => {
  try {
    if (mongoose.connection.readyState !== 1) {
      throw new Error('MongoDB connection is not established');
    }

    const totalSchoolsPromise = School.countDocuments().lean();
    const schoolNamesPromise = School.find({}, 'name').lean();
    const submissionStatsPromise = Report.aggregate([
      {
        $group: {
          _id: '$schoolName',
          submissionCount: { $sum: { $size: { $ifNull: ['$submissions', []] } } }
        }
      },
      {
        $project: {
          schoolName: '$_id',
          submissionCount: 1,
          _id: 0
        }
      },
      {
        $sort: { schoolName: 1 }
      }
    ])
      .allowDiskUse(true)
      .option({ maxTimeMS: 30000 });

    const [totalSchools, schools, submissionStats] = await Promise.all([
      totalSchoolsPromise,
      schoolNamesPromise,
      submissionStatsPromise
    ]);

    const schoolNames = schools.map(school => school.name);
    const totalSubmissions = submissionStats.reduce((sum, stat) => sum + stat.submissionCount, 0);
    const averageSubmissions = totalSchools > 0 ? (totalSubmissions / totalSchools).toFixed(2) : 0;

    res.status(200).json({
      success: true,
      data: {
        totalSchools,
        schoolNames,
        totalSubmissions,
        averageSubmissionsPerSchool: parseFloat(averageSubmissions),
        submissionsBySchool: submissionStats
      }
    });
  } catch (error) {
    console.error('Error fetching statistics:', error.message, error.stack);
    res.status(500).json({
      success: false,
      message: 'Failed to retrieve statistics',
      error: error.message
    });
  }
};

module.exports = { getStatistics };