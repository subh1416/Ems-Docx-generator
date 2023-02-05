const mongoose = require("mongoose");

const logsSchema = mongoose.Schema(
  {
    useremail: {
      type: String,
      required: true,
    },
    ip: {
      type: String,
      required: true,
    },
    urioriginal: {
      type: String,
      required: true,
    },
    uri: {
      type: String,
      required: true,
    },
    agent: {
      type: String,
      required: true,
    },
    referer: {
      type: String,
      required: true,
    },
  },
  { timestamps: true }
);

module.exports = mongoose.model("Logs", logsSchema, "Logs");
