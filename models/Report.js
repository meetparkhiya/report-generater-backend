const mongoose = require("mongoose");

const reportSchema = new mongoose.Schema(
    {
        employeeName: {
            type: String,
            required: true,
            trim: true
        },
        month: {
            type: String,
            required: true
        },
        year: {
            type: Number,
            required: true
        },
        report_file: {
            type: String,
            required: true
        },
        fileName: {
            type: String,
            required: true
        },
        fileSize: {
            type: Number
        }
    },
    { timestamps: true }
);

reportSchema.index({ employeeName: 1, month: 1, year: 1 });

module.exports = mongoose.model("Report", reportSchema);