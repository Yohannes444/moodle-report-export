const mongoose = require('mongoose');

const submissionSchema = new mongoose.Schema({
  courseId: { type: String, required: true },
  submissionName: { type: String, required: true },
  studentName: { type: String, default: 'Unknown' },
  studentUsername: { type: String, default: 'Unknown' },
  studentEmail: { type: String, default: 'Unknown' },
  dateSubmitted: { type: Date, required: true },
  directLink: { type: String, required: true }
});

const reportSchema = new mongoose.Schema({
  schoolName: {
    type: String,
    required: true,
    index: true
  },
  submissions: [submissionSchema], // Array of submission entries
  errorMessage: { type: String }, // For storing error details
  googleSheetsLink: {
    type: String,
    required: true
  },
  fileId: {
    type: String,
    required: true
  }
}, {
  timestamps: true
});

module.exports = mongoose.model('Report', reportSchema);