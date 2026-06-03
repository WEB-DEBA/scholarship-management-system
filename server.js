require("dotenv").config();

const dns = require("dns");
dns.setServers(["8.8.8.8", "8.8.4.4"]);

const express = require("express");
const bodyParser = require("body-parser");
const session = require("express-session");
const fs = require("fs");
const path = require("path");
const multer = require("multer");
const mongoose = require("mongoose");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");
const bcrypt = require("bcrypt");
const nodemailer = require("nodemailer");
const QRCode = require("qrcode");
const archiver = require("archiver");

const app = express();
const mailTransporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 587,
  secure: false,

  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
  },

  tls: {
    rejectUnauthorized: false
  }
});
const PORT = process.env.PORT || 3000;

app.set("view engine", "ejs");
app.use(express.static("public"));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.json());

app.use(session({
  secret: process.env.SESSION_SECRET || "scholarship_secret_key",
  resave: false,
  saveUninitialized: false
}));


const uploadDir = path.join(__dirname, "public", "uploads");

if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => {
    const fileName =
      Date.now() + "-" + Math.round(Math.random() * 1e9) + path.extname(file.originalname);
    cb(null, fileName);
  }
});

const upload = multer({ storage });

const studentSchema = new mongoose.Schema({
  registrationNo: { type: String, required: true, unique: true, index: true },
  aadhaar: { type: String, required: true, unique: true, index: true },
  accountNo: { type: String, index: true },
  phone: { type: String, index: true },

  name: String,
  fatherName: String,
  motherName: String,
  dob: String,
  gender: String,
  email: String,

  caste: String,
  casteCertificate: String,
  income: String,
  incomeCertificate: String,

  college: String,
  course: String,
  branch: String,
  semester: String,
  rollNo: String,
  lastExam: String,
  percentage: String,

  bankName: String,
  ifsc: String,
  accountHolder: String,

  village: String,
  post: String,
  district: String,
  state: String,
  pincode: String,

  aadhaarFile: String,
  casteFile: String,
  incomeFile: String,
  photoFile: String,
  bankFile: String,
  marksheetFile: String,

  status: { type: String, default: "Pending", index: true },
  paymentStatus: { type: String, default: "Pending", index: true },
  isDeleted: { type: Boolean, default: false, index: true },
  appliedDate: String
}, { timestamps: true });

const Student = mongoose.model("Student", studentSchema);

const SavedApplication = mongoose.model(
  "SavedApplication",
  new mongoose.Schema({}, { strict: false, timestamps: true })
);

const paymentSchema = new mongoose.Schema({
  id: { type: Number, unique: true, index: true },
  studentId: String,
  name: String,
  registrationNo: { type: String, index: true },
  rollNo: String,
  semester: String,
  branch: String,
  totalFee: String,
  amount: String,
  paymentDate: String,
  paymentMode: String,
  transactionId: String,
  status: { type: String, default: "Pending", index: true },
  remark: String,
  createdAtText: String,
  updatedAtText: String
}, { timestamps: true });

const Payment = mongoose.model("Payment", paymentSchema);

const noticeSchema = new mongoose.Schema({
  id: { type: Number, unique: true, index: true },
  title: String,
  message: String,
  date: String,
  category: { type: String, default: "Scholarship" },
  priority: { type: String, default: "High" }
}, { timestamps: true });

const Notice = mongoose.model("Notice", noticeSchema);

const Setting = mongoose.model("Setting", new mongoose.Schema({
  siteName: String,
  theme: String,
  noticeText: String
}, { timestamps: true }));

const ActivityLog = mongoose.model("ActivityLog", new mongoose.Schema({
  admin: String,
  action: String,
  studentId: String,
  details: String
}, { timestamps: true }));

const adminSchema = new mongoose.Schema({
  email: { type: String, unique: true, required: true },
  phone: { type: String, unique: true, required: true },
  password: { type: String, required: true },
  resetOtp: String,
  resetOtpExpire: Date
}, { timestamps: true });

const Admin = mongoose.model("Admin", adminSchema);

async function createDefaultAdmin() {
  const adminExists = await Admin.findOne();

  if (!adminExists) {
    const hashedPassword = await bcrypt.hash(
      process.env.ADMIN_PASSWORD || "pnssc/st.college",
      10
    );

    await Admin.create({
      email: process.env.ADMIN_EMAIL || "pnsscholarshipsystem@gmail.com",
      phone: process.env.ADMIN_PHONE || "7326032546",
      password: hashedPassword
    });

    console.log("✅ Default Admin Created");
  }
}

mongoose.connect(process.env.MONGO_URI || process.env.MONGO_URL)
  .then(async () => {
    console.log("✅ MongoDB Connected");
    await createDefaultAdmin();
  })
  .catch((err) => console.log("❌ MongoDB Error:", err));

mongoose.connection.once("open", () => {
  console.log("🟢 MongoDB Database Connected Successfully");
});

mongoose.connection.on("error", (err) => {
  console.log("🔴 MongoDB Error:", err);
});

function isAdmin(req, res, next) {
  if (req.session.admin) return next();
  res.redirect("/login");
}


async function addLog(action, studentId = "", details = "") {
  await ActivityLog.create({
    admin: "Admin",
    action,
    studentId,
    details
  });
}

app.get("/", async (req, res) => {
  const notices = await Notice.find().sort({ createdAt: -1 }).lean();
  res.render("index", { notices });
});

app.get("/notice", async (req, res) => {
  const notices = await Notice.find().sort({ createdAt: -1 }).lean();
  res.render("notice", { notices });
});

app.get("/apply", (req, res) => {
  res.render("apply", { error: null });
});
app.post("/apply", upload.fields([
  { name: "aadhaarFile", maxCount: 1 },
  { name: "casteFile", maxCount: 1 },
  { name: "incomeFile", maxCount: 1 },
  { name: "photoFile", maxCount: 1 },
  { name: "bankFile", maxCount: 1 },
  { name: "marksheetFile", maxCount: 1 }
]), async (req, res) => {
  try {
    const exists = await Student.findOne({
  $or: [
    { aadhaar: req.body.aadhaar },
    { registrationNo: req.body.registrationNo }
  ]
});

    if (exists) {
      return res.render("apply", {
        error: "⚠️ This Registration Number or Aadhaar Number is already registered."
      });
    }

    const newStudent = await Student.create({
      ...req.body,
      aadhaarFile: req.files.aadhaarFile?.[0]?.filename || "",
      casteFile: req.files.casteFile?.[0]?.filename || "",
      incomeFile: req.files.incomeFile?.[0]?.filename || "",
      photoFile: req.files.photoFile?.[0]?.filename || "",
      bankFile: req.files.bankFile?.[0]?.filename || "",
      marksheetFile: req.files.marksheetFile?.[0]?.filename || "",
      status: "Pending",
      paymentStatus: "Pending",
      isDeleted: false,
      appliedDate: new Date().toLocaleDateString()
    });

    try {
      await mailTransporter.sendMail({
        from: process.env.EMAIL_USER,
        to: process.env.ADMIN_EMAIL || "pnsscholarshipsystem@gmail.com",
        subject: "New Scholarship Application Submitted",
        html: `
          <h2>New Scholarship Application</h2>
          <p><b>Name:</b> ${newStudent.name || "N/A"}</p>
          <p><b>Registration No:</b> ${newStudent.registrationNo || "N/A"}</p>
          <p><b>Phone:</b> ${newStudent.phone || "N/A"}</p>
          <p><b>Branch:</b> ${newStudent.branch || "N/A"}</p>
          <p><b>Semester:</b> ${newStudent.semester || "N/A"}</p>
        `
      });
    } catch (mailError) {
      console.log("Apply email failed:", mailError.message);
    }

    res.redirect("/check");

  } catch (error) {

  console.log(error);

  if (error.code === 11000) {

    return res.render("apply", {
      error: "⚠️ This Registration Number or Aadhaar Number is already registered."
    });

  }

  res.render("apply", {
    error: "⚠️ Something went wrong. Please try again."
  });
}
});

app.get("/check", (req, res) => {
  res.render("check", {
    student: null,
    payment: null,
    searched: false
  });
});

app.post("/check", async (req, res) => {
  try {
    const searchValue = req.body.searchValue || req.body.aadhaar;

    const student = await Student.findOne({
      $or: [
        { aadhaar: searchValue },
        { registrationNo: searchValue }
      ],
      isDeleted: false
    }).lean();

    const payment = await Payment.findOne({
      registrationNo: searchValue
    }).lean();

    res.render("check", {
      student,
      payment,
      searched: true
    });
  } catch (error) {
    console.log(error);
    res.send("Search failed");
  }
});

app.get("/login", (req, res) => {
  res.render("login", { error: null });
});

app.post("/login", async (req, res) => {
  const { email, password } = req.body;

  const admin = await Admin.findOne({
    $or: [
      { email },
      { phone: email }
    ]
  });

  if (!admin) {
    return res.render("login", {
      error: "Invalid email/phone or password"
    });
  }

  const match = await bcrypt.compare(password, admin.password);

  if (!match) {
    return res.render("login", {
      error: "Invalid email/phone or password"
    });
  }

  req.session.admin = true;
  req.session.adminId = admin._id;

  res.redirect("/admin");
});

app.get("/forgot-password", (req, res) => {
  res.render("forgot-password", {
    error: null,
    success: null
  });
});
app.post("/forgot-password", async (req, res) => {
  try {

    const admin = await Admin.findOne({
      $or: [
        { email: req.body.email },
        { phone: req.body.email }
      ]
    });

    if (!admin) {
      return res.render("forgot-password", {
        error: "Admin not found",
        success: null
      });
    }

    const otp = Math.floor(
      100000 + Math.random() * 900000
    ).toString();

    admin.resetOtp = otp;
    admin.resetOtpExpire =
      Date.now() + 10 * 60 * 1000;

    await admin.save();

    await mailTransporter.sendMail({
      from: process.env.EMAIL_USER,
      to: admin.email,
      subject: "Scholarship Portal Password Reset OTP",

      html: `
        <h2>PNS Scholarship Portal</h2>
        <p>Your OTP is:</p>
        <h1>${otp}</h1>
        <p>Valid for 10 minutes.</p>
      `
    });

    res.render("reset-password", {
      error: null,
      email: req.body.email
    });

  } catch (error) {

  console.log("❌ EMAIL ERROR:");
  console.log(error);

  res.render("forgot-password", {
    error: error.message,
    success: null
  });

}
});
app.get("/student/receipt/:id", async (req, res) => {
  try {
    const student = await Student.findById(req.params.id).lean();

    if (!student) return res.send("Student not found");

    const verifyUrl =
      `${req.protocol}://${req.get("host")}/verify/${student._id}`;

    const qrDataUrl = await QRCode.toDataURL(verifyUrl);
    const qrImage = Buffer.from(
      qrDataUrl.replace(/^data:image\/png;base64,/, ""),
      "base64"
    );

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=${student.registrationNo}-receipt.pdf`
    );

    const doc = new PDFDocument({ margin: 50 });
    doc.pipe(res);

    doc.fontSize(22).text("PNS Scholarship Receipt", { align: "center" });
    doc.moveDown();

    doc.fontSize(12).text(`Application ID: ${student._id}`);
    doc.text(`Name: ${student.name}`);
    doc.text(`Registration No: ${student.registrationNo}`);
    doc.text(`Aadhaar: ${student.aadhaar}`);
    doc.text(`Phone: ${student.phone}`);
    doc.text(`College: ${student.college}`);
    doc.text(`Course: ${student.course}`);
    doc.text(`Branch: ${student.branch}`);
    doc.text(`Semester: ${student.semester}`);
    doc.text(`Status: ${student.status}`);
    doc.text(`Payment Status: ${student.paymentStatus}`);
    doc.text(`Applied Date: ${student.appliedDate || "N/A"}`);

    doc.moveDown();
    doc.text("Scan QR to verify this scholarship receipt:");
    doc.image(qrImage, { width: 120 });

    doc.end();

  } catch (error) {
    console.log(error);
    res.send("PDF receipt failed");
  }
});

app.get("/verify/:id", async (req, res) => {
  const student = await Student.findById(req.params.id).lean();

  if (!student) return res.send("Invalid QR Verification");

  res.render("verify", { student });
});
app.get("/admin/analytics", isAdmin, async (req, res) => {
  const students = await Student.find({ isDeleted: false }).lean();
  const payments = await Payment.find().lean();

  const analytics = {
    totalStudents: students.length,
    success: students.filter(s => s.status === "Success").length,
    pending: students.filter(s => s.status === "Pending").length,
    failed: students.filter(s => s.status === "Failed").length,
    totalPayments: payments.length,
    paidPayments: payments.filter(p => p.status === "Success").length
  };

  res.render("analytics", { analytics });
});
app.get("/admin/backup", isAdmin, async (req, res) => {
  try {
    const students = await Student.find().lean();
    const payments = await Payment.find().lean();
    const notices = await Notice.find().lean();
    const savedApplications = await SavedApplication.find().lean();
    const logs = await ActivityLog.find().lean();

    const backup = {
      students,
      payments,
      notices,
      savedApplications,
      logs,
      backupDate: new Date()
    };

    res.setHeader("Content-Type", "application/json");
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=scholarship-backup.json"
    );

    res.send(JSON.stringify(backup, null, 2));

  } catch (error) {
    console.log(error);
    res.send("Backup failed");
  }
});

app.post("/reset-password", async (req, res) => {
  const admin = await Admin.findOne({
    $or: [
      { email: req.body.email },
      { phone: req.body.email }
    ]
  });

  if (!admin) {
    return res.render("reset-password", {
      error: "Admin not found",
      email: req.body.email
    });
  }

  if (
    admin.resetOtp !== req.body.otp ||
    admin.resetOtpExpire < Date.now()
  ) {
    return res.render("reset-password", {
      error: "Invalid or expired OTP",
      email: req.body.email
    });
  }

  admin.password = await bcrypt.hash(req.body.password, 10);
  admin.resetOtp = null;
  admin.resetOtpExpire = null;

  await admin.save();

  res.redirect("/login");
});

app.get("/admin/profile", isAdmin, async (req, res) => {
  const admin = await Admin.findById(req.session.adminId).lean();

  res.render("admin-profile", {
    admin,
    error: null,
    success: null
  });
});

app.post("/admin/profile", isAdmin, async (req, res) => {
  const admin = await Admin.findById(req.session.adminId);

  admin.email = req.body.email;
  admin.phone = req.body.phone;

  if (req.body.password) {
    admin.password = await bcrypt.hash(req.body.password, 10);
  }

  await admin.save();

  res.render("admin-profile", {
    admin,
    error: null,
    success: "Profile updated successfully"
  });
});

app.get("/admin", isAdmin, async (req, res) => {
  try {
    const students = await Student.find({ isDeleted: false })
      .sort({ createdAt: -1 })
      .lean();

    const notices = await Notice.find().sort({ createdAt: -1 }).lean();
    const payments = await Payment.find().sort({ createdAt: -1 }).lean();

    res.render("admin", {
      students,
      notices,
      payments,
      successCount: students.filter(s => s.status === "Success").length,
      failedCount: students.filter(s => s.status === "Failed").length,
      pendingCount: students.filter(s => s.status === "Pending").length
    });
  } catch (error) {
    console.log(error);
    res.send("Admin dashboard failed");
  }
});

app.get("/admin/students", isAdmin, async (req, res) => {
  const students = await Student.find({ isDeleted: false })
    .sort({ createdAt: -1 })
    .lean();

  res.render("students", { students });
});

app.get("/admin/student/:id", isAdmin, async (req, res) => {
  try {
    const student = await Student.findById(req.params.id).lean();

    if (!student) {
      return res.send("Student not found");
    }

    const payments = await Payment.find({
      registrationNo: student.registrationNo
    }).sort({ createdAt: -1 }).lean();

    res.render("student-profile", {
      student,
      payments
    });
  } catch (error) {
    console.log(error);
    res.send("Student profile failed");
  }
});

app.post("/admin/status/:id", isAdmin, async (req, res) => {

  const student = await Student.findByIdAndUpdate(
    req.params.id,
    {
      status: req.body.status
    },
    { new: true }
  );

  if (student && student.email) {

    let subject = "";
    let html = "";

    if (req.body.status === "Success") {

      subject = "Scholarship Application Approved";

      html = `
        <h2 style="color:green;">🎉 Scholarship Approved</h2>

        <p>Dear ${student.name},</p>

        <p>Your scholarship application has been approved successfully.</p>

        <p><b>Registration No:</b> ${student.registrationNo}</p>
        <p><b>Status:</b> Success</p>

        <p>Thank you for using PNS Scholarship Portal.</p>
      `;

    } else if (req.body.status === "Failed") {

      subject = "Scholarship Application Rejected";

      html = `
        <h2 style="color:red;">❌ Scholarship Rejected</h2>

        <p>Dear ${student.name},</p>

        <p>Your scholarship application could not be approved.</p>

        <p><b>Registration No:</b> ${student.registrationNo}</p>
        <p><b>Status:</b> Failed</p>
      `;

    } else {

      subject = "Scholarship Application Under Review";

      html = `
        <h2 style="color:orange;">⏳ Application Under Review</h2>

        <p>Dear ${student.name},</p>

        <p>Your scholarship application is currently under review.</p>

        <p><b>Registration No:</b> ${student.registrationNo}</p>
        <p><b>Status:</b> Pending</p>
      `;
    }

    try {

      await mailTransporter.sendMail({
        from: process.env.EMAIL_USER,
        to: student.email,
        subject,
        html
      });

      console.log("✅ Status Email Sent");

    } catch (err) {

      console.log("❌ Email Error:", err.message);

    }
  }

  await addLog("Status Updated", req.params.id, req.body.status);

  res.redirect("/admin");
});


app.post("/admin/delete/:id", isAdmin, async (req, res) => {
  await Student.findByIdAndUpdate(req.params.id, {
    isDeleted: true
  });

  await addLog("Student Soft Deleted", req.params.id, "Moved to deleted records");

  res.redirect("/admin");
});
app.post("/admin/application/save/:id", isAdmin, async (req, res) => {
  try {
    const student = await Student.findById(req.params.id).lean();

    if (!student) {
      return res.send("Application not found");
    }

    await SavedApplication.create({
      ...student,
      originalStudentId: student._id,
      savedAt: new Date().toLocaleString()
    });

    await Student.findByIdAndUpdate(req.params.id, {
      isDeleted: true
    });

    await addLog("Application Saved", req.params.id, "Saved permanent application");

    if (student.email) {
      try {
        await mailTransporter.sendMail({
          from: process.env.EMAIL_USER,
          to: student.email,
          subject: "Scholarship Application Approved",
          html: `
            <h2 style="color:green;">🎉 Application Approved</h2>
            <p>Dear ${student.name},</p>
            <p>Your scholarship application has been approved successfully by the administration.</p>
            <p><b>Registration No:</b> ${student.registrationNo}</p>
            <p><b>Branch:</b> ${student.branch}</p>
            <p><b>Semester:</b> ${student.semester}</p>
            <p>You can now check your application status on the PNS Scholarship Portal.</p>
            <br>
            <p>Regards,<br>PNS Scholarship Portal</p>
          `
        });

        console.log("✅ Approval Email Sent");
      } catch (err) {
        console.log("❌ Approval Email Error:", err.message);
      }
    }

    res.redirect("/admin");
  } catch (error) {
    console.log(error);
    res.send("Application save failed");
  }
});

app.get("/admin/application-details", isAdmin, async (req, res) => {
  const applications = await SavedApplication.find()
    .sort({ createdAt: -1 })
    .lean();

  res.render("application-details", { applications });
});

app.get("/admin/saved-application/:id", isAdmin, async (req, res) => {
  const student = await SavedApplication.findById(req.params.id).lean();

  if (!student) {
    return res.send("Saved application not found");
  }

  res.render("student-profile", {
    student,
    payments: []
  });
});

app.get("/admin/application-details/edit/:id", isAdmin, async (req, res) => {
  const application = await SavedApplication.findById(req.params.id).lean();

  if (!application) {
    return res.send("Saved application not found");
  }

  res.render("edit-applications", { application });
});

app.post("/admin/application-details/edit/:id", isAdmin, async (req, res) => {
  await SavedApplication.findByIdAndUpdate(req.params.id, {
    ...req.body,
    updatedAtText: new Date().toLocaleString()
  });

  await addLog("Saved Application Edited", req.params.id, "Edited saved application");

  res.redirect("/admin/application-details");
});

app.post("/admin/application-details/delete/:id", isAdmin, async (req, res) => {
  await SavedApplication.findByIdAndDelete(req.params.id);

  await addLog("Saved Application Deleted", req.params.id, "Deleted saved application");

  res.redirect("/admin/application-details");
});


app.get("/admin/payments", isAdmin, async (req, res) => {
  const payments = await Payment.find().sort({ createdAt: -1 }).lean();
  res.render("payments", { payments });
});

app.post("/admin/payments", isAdmin, async (req, res) => {
  const payment = await Payment.create({
    id: Date.now(),
    studentId: req.body.studentId || "",
    name: req.body.name,
    registrationNo: req.body.registrationNo,
    rollNo: req.body.rollNo,
    semester: req.body.semester,
branch: req.body.branch,
totalFee: req.body.totalFee,
amount: req.body.amount,
paymentDate: req.body.paymentDate,
    paymentMode: req.body.paymentMode,
    transactionId: req.body.transactionId,
    status:
  Number(req.body.amount || 0) >= Number(req.body.totalFee || 0)
    ? "Success"
    : "Pending",
    remark: req.body.remark,
    createdAtText: new Date().toLocaleString()
  });

  if (req.body.studentId && req.body.status === "Success") {
  await Student.findByIdAndUpdate(req.body.studentId, {
    status: "Success",
    paymentStatus: "Paid"
  });
}

  await addLog("Payment Added", req.body.studentId || "", payment.registrationNo || "");

  res.redirect("/admin/payments");
});

app.get("/admin/payments/edit/:id", isAdmin, async (req, res) => {
  const payment = await Payment.findOne({ id: Number(req.params.id) }).lean();

  if (!payment) {
    return res.send("Payment record not found");
  }

  res.render("edit-payment", { payment });
});
app.post("/admin/payments/edit/:id", isAdmin, async (req, res) => {
  const updatedPayment = await Payment.findOneAndUpdate(
    { id: Number(req.params.id) },
    {
      name: req.body.name,
      registrationNo: req.body.registrationNo,
      rollNo: req.body.rollNo,
      semester: req.body.semester,
      branch: req.body.branch,
      totalFee: req.body.totalFee,
      amount: req.body.amount,
      paymentDate: req.body.paymentDate,
      paymentMode: req.body.paymentMode,
      transactionId: req.body.transactionId,
      status:
        Number(req.body.amount || 0) >= Number(req.body.totalFee || 0)
          ? "Success"
          : "Pending",
      remark: req.body.remark,
      updatedAtText: new Date().toLocaleString()
    },
    { returnDocument: "after" }
  );

  if (updatedPayment && updatedPayment.registrationNo) {
    const student = await Student.findOneAndUpdate(
      { registrationNo: updatedPayment.registrationNo },
      {
        status: updatedPayment.status === "Success" ? "Success" : "Pending",
        paymentStatus: updatedPayment.status === "Success" ? "Paid" : "Pending"
      },
      { returnDocument: "after" }
    );

    if (student && student.email) {

  let subject = "";
  let html = "";

  if (updatedPayment.status === "Success") {

    subject = "Scholarship Payment Successful";

    html = `
      <h2>PNS Scholarship Portal</h2>
      <p>Dear ${student.name},</p>

      <p>Your scholarship payment has been successfully completed.</p>

      <p><b>Total Fee:</b> ₹${updatedPayment.totalFee}</p>
      <p><b>Paid Amount:</b> ₹${updatedPayment.amount}</p>
      <p><b>Status:</b> Success</p>

      <p>Thank you.</p>
    `;

  } else if (updatedPayment.status === "Pending") {

    const remaining =
      Number(updatedPayment.totalFee || 0) -
      Number(updatedPayment.amount || 0);

    subject = "Scholarship Payment Pending";

    html = `
      <h2>PNS Scholarship Portal</h2>
      <p>Dear ${student.name},</p>

      <p>Your payment record has been updated.</p>

      <p><b>Total Fee:</b> ₹${updatedPayment.totalFee}</p>
      <p><b>Paid Amount:</b> ₹${updatedPayment.amount}</p>
      <p><b>Remaining Amount:</b> ₹${remaining}</p>

      <p><b>Status:</b> Pending</p>
    `;
  }

  await mailTransporter.sendMail({
    from: process.env.EMAIL_USER,
    to: student.email,
    subject,
    html
  });
}
  }

  await addLog("Payment Edited", "", req.params.id);

  res.redirect("/admin/payments");
});

app.post("/admin/payments/delete/:id", isAdmin, async (req, res) => {
  await Payment.findOneAndDelete({ id: Number(req.params.id) });

  await addLog("Payment Deleted", "", req.params.id);

  res.redirect("/admin/payments");
});

app.post("/admin/whatsapp", isAdmin, (req, res) => {
  const phones = req.body.phones;
  const message = encodeURIComponent(req.body.message || "");

  if (!phones) return res.redirect("/admin");

  const phoneList = Array.isArray(phones) ? phones : [phones];
  const firstPhone = phoneList[0];

  res.redirect(`https://wa.me/91${firstPhone}?text=${message}`);
});

app.post("/admin/direct-whatsapp", isAdmin, (req, res) => {
  const phone = req.body.phone;
  const message = encodeURIComponent(req.body.message || "");

  if (!phone) return res.redirect("/admin");

  res.redirect(`https://wa.me/91${phone}?text=${message}`);
});

app.get("/admin/fraud", isAdmin, async (req, res) => {
  const students = await Student.find({ isDeleted: false }).lean();

  const aadhaarMap = {};
  const bankMap = {};

  students.forEach(s => {
    if (s.aadhaar) aadhaarMap[s.aadhaar] = (aadhaarMap[s.aadhaar] || 0) + 1;
    if (s.accountNo) bankMap[s.accountNo] = (bankMap[s.accountNo] || 0) + 1;
  });

  const duplicateAadhaar = students.filter(s => aadhaarMap[s.aadhaar] > 1);
  const duplicateBank = students.filter(s => bankMap[s.accountNo] > 1);

  res.render("fraud", {
    duplicateAadhaar,
    duplicateBank
  });
});

app.get("/admin/settings", isAdmin, async (req, res) => {
  let settings = await Setting.findOne().lean();

  if (!settings) {
    settings = await Setting.create({
      siteName: "PNS Scholarship Portal",
      theme: "Government Digital Portal",
      noticeText: "SC/ST Scholarship Application Open"
    });
  }

  res.render("settings", { settings });
});

app.post("/admin/settings", isAdmin, async (req, res) => {
  let settings = await Setting.findOne();

  if (!settings) {
    settings = new Setting();
  }

  settings.siteName = req.body.siteName;
  settings.theme = req.body.theme;
  settings.noticeText = req.body.noticeText;

  await settings.save();

  await addLog("Settings Updated", "", "Portal settings updated");

  res.redirect("/admin/settings");
});

app.get("/admin/export/excel", isAdmin, async (req, res) => {
  const students = await Student.find({ isDeleted: false }).lean();

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Students");

  sheet.columns = [
    { header: "Name", key: "name", width: 25 },
    { header: "Registration No", key: "registrationNo", width: 20 },
    { header: "Aadhaar", key: "aadhaar", width: 20 },
    { header: "Phone", key: "phone", width: 15 },
    { header: "College", key: "college", width: 25 },
    { header: "Course", key: "course", width: 20 },
    { header: "Branch", key: "branch", width: 20 },
    { header: "Semester", key: "semester", width: 15 },
    { header: "Bank Account", key: "accountNo", width: 22 },
    { header: "IFSC", key: "ifsc", width: 16 },
    { header: "Payment", key: "paymentStatus", width: 16 },
    { header: "Status", key: "status", width: 15 }
  ];

  students.forEach(s => sheet.addRow(s));

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  res.setHeader(
    "Content-Disposition",
    "attachment; filename=students.xlsx"
  );

  await workbook.xlsx.write(res);
  res.end();
});
app.get("/admin/export/pdf", isAdmin, async (req, res) => {
  const students = await Student.find({ isDeleted: false }).lean();

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", "attachment; filename=students.pdf");

  const doc = new PDFDocument({ margin: 40 });
  doc.pipe(res);

  doc.fontSize(22).text("PNS Scholarship Student Report", {
    align: "center"
  });

  doc.moveDown();

  students.forEach((s, i) => {
    doc.fontSize(11).text(
      `${i + 1}. ${s.name || "N/A"} | Reg: ${s.registrationNo || "N/A"} | Aadhaar: ${s.aadhaar || "N/A"} | Branch: ${s.branch || "N/A"} | Sem: ${s.semester || "N/A"} | Payment: ${s.paymentStatus || "Pending"} | Status: ${s.status || "Pending"}`
    );
  });

  doc.end();
});

app.post("/admin/notice", isAdmin, async (req, res) => {
  await Notice.create({
    id: Date.now(),
    title: req.body.title,
    message: req.body.message,
    date: new Date().toLocaleDateString(),
    category: "Scholarship",
    priority: "High"
  });

  await addLog("Notice Added", "", req.body.title);

  res.redirect("/admin");
});

app.post("/admin/notice/delete/:index", isAdmin, async (req, res) => {
  const notices = await Notice.find().sort({ createdAt: -1 }).lean();
  const notice = notices[Number(req.params.index)];

  if (notice) {
    await Notice.findByIdAndDelete(notice._id);
    await addLog("Notice Deleted", "", notice.title);
  }

  res.redirect("/admin");
});

app.get("/admin/logs", isAdmin, async (req, res) => {
  const logs = await ActivityLog.find()
    .sort({ createdAt: -1 })
    .limit(100)
    .lean();

  res.render("logs", { logs });
});
app.get("/admin/payment-receipt/:id", isAdmin, async (req, res) => {
  try {
    const payment = await Payment.findOne({
      id: Number(req.params.id)
    }).lean();

    if (!payment) {
      return res.send("Payment record not found");
    }

    const remaining =
      Number(payment.totalFee || 0) - Number(payment.amount || 0);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader(
      "Content-Disposition",
      `inline; filename=receipt-${payment.registrationNo}.pdf`
    );

    const doc = new PDFDocument({ margin: 50 });
    doc.pipe(res);

    doc.rect(25, 25, 545, 740).stroke();

    doc.fontSize(22).text("PNS SCHOLARSHIP PORTAL", 50, 45, {
      align: "center"
    });

    doc.fontSize(10).text("SC/ST Scholarship Management System", {
      align: "center"
    });

    doc.fontSize(16).text("MONEY RECEIPT", 50, 95, {
      align: "center",
      underline: true
    });

    doc.fontSize(10)
      .text(`Receipt No: RCPT-${payment.id}`, 50, 130)
      .text(`Date: ${payment.paymentDate || "N/A"}`, 400, 130);

    doc.moveTo(50, 155).lineTo(545, 155).stroke();

    doc.fontSize(12);
    doc.text(`Received from: ${payment.name || "N/A"}`, 50, 180);
    doc.text(`Registration No: ${payment.registrationNo || "N/A"}`, 50, 205);
    doc.text(`Roll No: ${payment.rollNo || "N/A"}`, 50, 230);
    doc.text(`Semester: ${payment.semester || "N/A"}`, 330, 230);
    doc.text(`Branch: ${payment.branch || "N/A"}`, 50, 255);

    doc.moveTo(50, 290).lineTo(545, 290).stroke();

    doc.fontSize(13).text("Payment Details", 50, 310);

    doc.fontSize(12);
    doc.text(`Total Course Fee: ₹${payment.totalFee || 0}`, 70, 345);
    doc.text(`Paid Amount: ₹${payment.amount || 0}`, 70, 370);
    doc.text(`Remaining Amount: ₹${remaining}`, 70, 395);
    doc.text(`Payment Mode: ${payment.paymentMode || "N/A"}`, 70, 420);
    doc.text(`Money Receipt No: PNS-${new Date().getFullYear()}-${payment.id}`, 70, 445);
    doc.text(`Payment Status: ${payment.status || "Pending"}`, 70, 470);

    doc.moveTo(50, 510).lineTo(545, 510).stroke();

    doc.fontSize(11).text(
      "This is a system generated money receipt for scholarship payment record.",
      50,
      535,
      { align: "center" }
    );

    doc.fontSize(11).text("Student Signature", 70, 650);
    doc.fontSize(11).text("Authorized Signature", 390, 650);

    doc.moveTo(70, 635).lineTo(190, 635).stroke();
    doc.moveTo(390, 635).lineTo(520, 635).stroke();

    doc.fontSize(9).text("Generated by PNS Scholarship Portal", 50, 720, {
      align: "center"
    });

    doc.end();

  } catch (error) {
    console.log(error);
    res.send("Receipt generation failed");
  }
});

app.get("/logout", (req, res) => {
  req.session.destroy(() => {
    res.redirect("/");
  });
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});