const mongoose = require('mongoose');

const schoolSchema = new mongoose.Schema({
  name: {
    type: String,
    required: true,
    trim: true,
  },
  baseUrl: {
    type: String,
    required: true,
    trim: true,
  },
  token: {
    type: String,
    required: true,
  },
  serviceAccountKey: {
    type: String,
    default: './moodleapi-456414-245626e993c8.json',
  }
}, {
  timestamps: true
});

const School = mongoose.model('School', schoolSchema);

module.exports = School;
