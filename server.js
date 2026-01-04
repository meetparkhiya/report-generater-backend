import express from "express";
import multer from "multer";
import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import cors from "cors";
import dotenv from "dotenv";
import connectDB from "./config/db.js";
import Report from "./models/Report.js";


dotenv.config();

const app = express();
const upload = multer({ dest: "uploads/" });



app.use(cors());
app.use(express.json());

app.use("/uploads", express.static(path.join(process.cwd(), "uploads")));

connectDB();

if (!fs.existsSync("uploads")) {
  fs.mkdirSync("uploads");
}

// 📝 GENERATE WORD FROM EXCEL DATA
app.post("/generate-word-from-excel", upload.single("template"), async (req, res) => {
  let templateFile = null;

  try {
    templateFile = "tasks.docx";

    if (!fs.existsSync(templateFile)) {
      return res.status(400).json({
        error: "Template file not found",
        message: "Please ensure tasks.docx template exists in the project root"
      });
    }

    const excelData = JSON.parse(req.body.data);
    const employeeName = req.body.employeeName;
    const month = req.body.month || "";
    const year = parseInt(req.body.year) || new Date().getFullYear();
    const generatedDate = req.body.generatedDate || "";

    const content = fs.readFileSync(templateFile, "binary");
    const zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => "",
      delimiters: { start: "{{", end: "}}" },
    });

    const wordTemplateData = {
      employeeName: employeeName,
      month: month,
      year: year,
      generatedDate: generatedDate,
      ...excelData
    };

    doc.setData(wordTemplateData);
    doc.render();

    const buf = doc.getZip().generate({
      type: "nodebuffer",
      compression: "DEFLATE"
    });

    // Create employee folder structure: uploads/EmployeeName/Month_Year/
    const employeeFolderName = employeeName.replace(/\s+/g, "_");
    const monthFolderName = `${month}_${year}`;

    const employeeFolderPath = path.join("uploads", employeeFolderName);
    const monthFolderPath = path.join(employeeFolderPath, monthFolderName);

    // Create folders if they don't exist
    if (!fs.existsSync(employeeFolderPath)) {
      fs.mkdirSync(employeeFolderPath, { recursive: true });
      console.log(`📁 Created employee folder: ${employeeFolderPath}`);
    }

    if (!fs.existsSync(monthFolderPath)) {
      fs.mkdirSync(monthFolderPath, { recursive: true });
      console.log(`📁 Created month folder: ${monthFolderPath}`);
    }

    // ✅ CHECK IF OLD REPORT EXISTS IN DATABASE
    const existingReport = await Report.findOne({
      employeeName: employeeName,
      month: month,
      year: year
    });

    if (existingReport) {
      // Delete old file from disk
      if (fs.existsSync(existingReport.report_file)) {
        fs.unlinkSync(existingReport.report_file);
        console.log(`🗑️ Deleted old file: ${existingReport.report_file}`);
      }

      // Delete old record from database
      // await Report.deleteOne({ _id: existingReport._id });
      console.log(`🗑️ Deleted old record from database`);
    }

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-").split("T")[0];
    const filename = `${employeeFolderName}_${month}_${year}_${timestamp}.docx`;
    const filepath = path.join(monthFolderPath, filename);

    // Save file to disk
    fs.writeFileSync(filepath, buf);
    console.log(`✅ Document saved: ${filepath}`);

    // Get file size
    const stats = fs.statSync(filepath);
    const fileSizeInBytes = stats.size;

    // ✅ SAVE TO DATABASE
    const newReport = new Report({
      employeeName: employeeName,
      month: month,
      year: year,
      report_file: filepath,
      fileName: filename,
      fileSize: fileSizeInBytes
    });

    await newReport.save();
    console.log(`💾 Report saved to database: ${newReport._id}`);

    // Send response to client
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=${employeeFolderName}_${month}_${year}_report.docx`
    );

    res.send(buf);

  } catch (err) {
    console.error("❌ ERROR:", err);

    if (err.properties && err.properties.errors) {
      console.error("📋 Template Errors:");
      err.properties.errors.forEach((error, index) => {
        console.error(`\n  Error ${index + 1}:`);
        console.error(`  - Message: ${error.message}`);
        console.error(`  - Tag: ${error.properties?.xtag}`);
        console.error(`  - Explanation: ${error.properties?.explanation}`);
      });

      res.status(400).json({
        error: "Template Error",
        message: "Word template has formatting issues",
        details: err.properties.errors.map(e => ({
          type: e.properties?.id,
          tag: e.properties?.xtag,
          issue: e.message
        }))
      });
    } else {
      res.status(500).json({
        error: "Server Error",
        message: err.message || "Failed to generate document"
      });
    }
  }
});

// 📜 GET ALL REPORTS (HISTORY)
app.get("/reports", async (req, res) => {
  try {
    const { employeeName, month, year, page = 1, limit = 20 } = req.query;

    const query = {};
    if (employeeName) query.employeeName = new RegExp(employeeName, 'i');
    if (month) query.month = month;
    if (year) query.year = parseInt(year);

    const skip = (parseInt(page) - 1) * parseInt(limit);

    const reports = await Report.find(query)
      .sort({ createdAt: -1 })
      .limit(parseInt(limit))
      .skip(skip)
      .lean();

    const total = await Report.countDocuments(query);

    res.json({
      success: true,
      count: reports.length,
      total: total,
      page: parseInt(page),
      totalPages: Math.ceil(total / parseInt(limit)),
      reports: reports
    });
  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message
    });
  }
});

// 📥 DOWNLOAD REPORT BY ID
app.get("/reports/download/:id", async (req, res) => {
  try {
    const report = await Report.findById(req.params.id);

    if (!report) {
      return res.status(404).json({
        success: false,
        message: "Report not found"
      });
    }

    if (!fs.existsSync(report.report_file)) {
      return res.status(404).json({
        success: false,
        message: "File not found on server"
      });
    }

    res.download(report.report_file, report.fileName);
  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message
    });
  }
});

// 🗑️ DELETE REPORT
app.delete("/reports/:id", async (req, res) => {
  try {
    const report = await Report.findById(req.params.id);

    if (!report) {
      return res.status(404).json({
        success: false,
        message: "Report not found"
      });
    }

    // Delete file from disk
    if (fs.existsSync(report.report_file)) {
      fs.unlinkSync(report.report_file);
      console.log(`🗑️ Deleted file: ${report.report_file}`);
    }

    // Delete from database
    await Report.deleteOne({ _id: report._id });

    res.json({
      success: true,
      message: "Report deleted successfully"
    });
  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message
    });
  }
});

// 📊 GET STATISTICS
app.get("/reports/stats", async (req, res) => {
  try {
    const totalReports = await Report.countDocuments();
    const uniqueEmployees = await Report.distinct("employeeName");

    const recentReports = await Report.find()
      .sort({ createdAt: -1 })
      .limit(5)
      .select('employeeName month year createdAt')
      .lean();

    const employeeStats = await Report.aggregate([
      {
        $group: {
          _id: "$employeeName",
          totalReports: { $sum: 1 },
          lastGenerated: { $max: "$createdAt" }
        }
      },
      { $sort: { totalReports: -1 } }
    ]);

    res.json({
      success: true,
      statistics: {
        totalReports: totalReports,
        totalEmployees: uniqueEmployees.length,
        recentReports: recentReports,
        employeeStats: employeeStats
      }
    });
  } catch (err) {
    res.status(500).json({
      success: false,
      error: err.message
    });
  }
});

// 🔍 INSPECT TEMPLATE
app.post("/inspect-template", upload.single("template"), (req, res) => {
  let templateFile = null;

  try {
    templateFile = req.file ? req.file.path : "tasks.docx";

    const content = fs.readFileSync(templateFile, "binary");
    const zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    const tags = doc.getTags();
    const fullText = doc.getFullText();

    res.json({
      success: true,
      tags: tags,
      preview: fullText.substring(0, 500),
      tagCount: Object.keys(tags).length
    });

    if (req.file) {
      fs.unlinkSync(templateFile);
    }
  } catch (err) {
    console.error("Inspection error:", err);
    res.status(400).json({
      success: false,
      error: err.message,
      details: err.properties
    });

    if (templateFile && req.file && fs.existsSync(templateFile)) {
      fs.unlinkSync(templateFile);
    }
  }
});

app.get("/health", (req, res) => {
  res.json({
    status: "ok",
    message: "Excel to Word Generator Server is running",
    mongodb: "Connected via existing config",
    endpoints: {
      generate: "POST /generate-word-from-excel",
      reports: "GET /reports",
      download: "GET /reports/download/:id",
      delete: "DELETE /reports/:id",
      stats: "GET /reports/stats",
      inspect: "POST /inspect-template",
      health: "GET /health"
    }
  });
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`\n✅ Server running on http://localhost:${PORT}`);
});