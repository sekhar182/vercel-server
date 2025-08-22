import dotenv from "dotenv";
import express from "express";
import cors from "cors";
import nodemailer from "nodemailer";
import fs from "fs";
import XLSX from "xlsx";
import process from "process";

// 'process' is always available in Node.js; no need for fallback in ES modules.

dotenv.config();

// Initialize Express App
const app = express();

// Middleware
app.use(express.json());
app.use(cors());

// Path to Excel file
const FILE_PATH = "data.xlsx";
const SHEET_NAME = "Form Submissions";

// Function to Auto-Adjust Column Widths
const autoAdjustColumnWidths = (worksheet, data) => {
  const columnWidths = Object.keys(data[0] || {}).map((key) => ({
    wch: Math.max(
      key.length, // Column header width
      ...data.map((row) => String(row[key] || "").length) // Max content width
    ) + 2, // Add some padding
  }));
  worksheet["!cols"] = columnWidths;
};

// Function to Append Data to Excel
const appendToExcel = (data) => {
  let workbook;
  let worksheet;
  let existingData = [];

  // Check if file exists
  if (fs.existsSync(FILE_PATH)) {
    workbook = XLSX.readFile(FILE_PATH);
    worksheet = workbook.Sheets[SHEET_NAME];

    // If sheet exists, load existing data
    if (worksheet) {
      existingData = XLSX.utils.sheet_to_json(worksheet) || [];
    } else {
      worksheet = XLSX.utils.json_to_sheet([]);
      XLSX.utils.book_append_sheet(workbook, worksheet, SHEET_NAME);
    }
  } else {
    // Create a new workbook and worksheet
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.json_to_sheet([]);
    XLSX.utils.book_append_sheet(workbook, worksheet, SHEET_NAME);
  }

  // Append new data
  existingData.push(data);

  // Convert back to worksheet
  const updatedWorksheet = XLSX.utils.json_to_sheet(existingData);

  // Auto-adjust column widths
  autoAdjustColumnWidths(updatedWorksheet, existingData);

  // Save updated worksheet to the workbook
  workbook.Sheets[SHEET_NAME] = updatedWorksheet;
  XLSX.writeFile(workbook, FILE_PATH);
  console.log("✅ Data Saved to Excel Successfully (with readable format)");
};

// POST Route to Handle Form Submission
app.post("/send-email", async (req, res) => {
  console.log("Received Data:", req.body);

  const { fullName, email, phone, subject, message, preferredContact } = req.body;

  try {
    // Save to Excel
    appendToExcel({
      fullName,
      email,
      phone,
      subject,
      message,
      preferredContact,
      date: new Date().toLocaleString(),
    });

    // Send Email
    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: process.env.EMAIL_USER, // Your email
        pass: process.env.EMAIL_PASS, // Your email password or app password
      },
    });

    const mailOptions = {
      from: process.env.EMAIL_USER,
      to: email, // Send to user
      subject: `Thank you for contacting us, ${fullName}!`,
      text: `Hello ${fullName},\n\nThank you for reaching out. We will get back to you soon.\n\nYour Message:\n${message}\n\nBest regards,\nTeam_NCS`,
    };

    await transporter.sendMail(mailOptions);
    console.log("✅ Email Sent Successfully");

    res.status(200).json({ message: "Your message sent successfully!" });
  } catch (error) {
    console.error("❌ Error:", error);
    res.status(500).json({ message: "Failed to process request." });
  }
});

// Start Server
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
