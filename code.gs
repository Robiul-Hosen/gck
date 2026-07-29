// =====================================================================
// স্কুল ও শিক্ষার্থী ব্যবস্থাপনা সিস্টেম — Apps Script Backend (সম্পূর্ণ, v3)
// =====================================================================

const SUPER_ADMIN_CONTACT = { name: 'Robiul Hosen Sojib', phone: '01617703090' };
const RESTRICTED_MSG = 'Teacher বা Student-এর তথ্য সংশোধন/মুছে ফেলতে Super Admin-এর সাথে যোগাযোগ করুন।';
const SUPER_ADMIN_ONLY_MSG = 'এই কাজটি শুধুমাত্র Super Admin করতে পারবেন।';
const ADMIN_ONLY_MSG = 'স্কুল নিবন্ধন অনুমোদন/বাতিল/ফেরত পাঠানো শুধুমাত্র Admin বা Super Admin করতে পারবেন।';

// প্রতিটা শীটের সঠিক হেডার — কোনো শীট মিসিং থাকলে getSheet() নিজে থেকেই এই হেডার দিয়ে তৈরি করে নেবে
// (নতুন ফিচারের জন্য নতুন শীট ম্যানুয়ালি বানাতে ভুলে গেলেও সিস্টেম ক্র্যাশ করবে না)
const SHEET_SCHEMAS = {
  'Schools': ['school_id', 'school_name', 'email', 'password_hash', 'phone', 'status', 'created_date', 'admin_note'],
  'Users': ['user_id', 'name', 'email', 'password_hash', 'role', 'school_id', 'created_date', 'status'],
  'Classes': ['class_id', 'class_name'],
  'Students': ['student_id', 'name', 'class_id', 'school_id', 'roll_no', 'contact', 'father_name', 'mother_name'],
  'Teachers': ['teacher_id', 'name', 'subject', 'class_id', 'school_id', 'contact'],
  'Exams': ['exam_id', 'exam_name', 'exam_date', 'status', 'fee_per_student'],
  'Rooms': ['room_id', 'exam_id', 'room_name', 'capacity'],
  'SeatPlan': ['seat_id', 'exam_id', 'room_id', 'student_id', 'student_name', 'school_id', 'roll_no'],
  'Invigilators': ['assign_id', 'exam_id', 'room_id', 'teacher_id', 'teacher_name', 'teacher_school_id', 'teacher_phone'],
  'Payments': ['payment_id', 'exam_id', 'school_id', 'amount', 'method', 'status', 'payment_date', 'note'],
  'Config': ['key', 'value'],
  'AuditLog': ['log_id', 'timestamp', 'actor_role', 'actor_name', 'action', 'target_school', 'note']
};

// ===== হেল্পার =====
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = SHEET_SCHEMAS[name];
    if (headers) sheet.appendRow(headers);
  }
  return sheet;
}
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}
function generateId(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}
function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}
function toBanglaDigits(input) {
  const map = { '0':'০','1':'১','2':'২','3':'৩','4':'৪','5':'৫','6':'৬','7':'৭','8':'৮','9':'৯' };
  return String(input).replace(/[0-9]/g, d => map[d]);
}
function isSuperAdmin(data) { return data.requesterRole === 'super_admin'; }
function isAdmin(data) { return data.requesterRole === 'admin'; }
function isAdminOrSuper(data) { return data.requesterRole === 'admin' || data.requesterRole === 'super_admin'; }

function testHash() {
  Logger.log(hashPassword("admin01"));
}

// =====================================================================
// doPost রাউটার
// =====================================================================
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  let result;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    switch (action) {
      case 'registerSchool': result = registerSchool(data); break;
      case 'login': result = login(data); break;

      case 'getAllSchools': result = getAllSchools(); break;
      case 'getSchoolDetail': result = getSchoolDetail(data); break;
      case 'approveSchool': result = approveSchool(data); break;
      case 'rejectSchool': result = rejectSchool(data); break;
      case 'returnForCorrection': result = returnForCorrection(data); break;
      case 'updateSchool': result = updateSchool(data); break;
      case 'deleteSchool': result = deleteSchool(data); break;
      case 'resetSchoolPassword': result = resetSchoolPassword(data); break;

      case 'addAdmin': result = addAdmin(data); break;
      case 'getAdmins': result = getAdmins(); break;
      case 'updateAdminStatus': result = updateAdminStatus(data); break;
      case 'resetAdminPassword': result = resetAdminPassword(data); break;

      case 'getAuditLog': result = getAuditLog(data); break;

      case 'getStudentRegistrationStatus': result = getStudentRegistrationStatus(); break;
      case 'setStudentRegistrationStatus': result = setStudentRegistrationStatus(data); break;

      case 'addClass': result = addClass(data); break;
      case 'getClasses': result = getClasses(); break;

      case 'addStudent': result = addStudent(data); break;
      case 'getStudents': result = getStudents(data); break;
      case 'updateStudent': result = updateStudent(data); break;
      case 'deleteStudent': result = deleteStudent(data); break;
      case 'updateStudentRoll': result = updateStudentRoll(data); break;
      case 'getAllStudentsGrouped': result = getAllStudentsGrouped(); break;
      case 'getStudentsForRollAssignment': result = getStudentsForRollAssignment(); break;

      case 'addTeacher': result = addTeacher(data); break;
      case 'getTeachers': result = getTeachers(data); break;
      case 'updateTeacher': result = updateTeacher(data); break;
      case 'deleteTeacher': result = deleteTeacher(data); break;
      case 'getAllTeachersGrouped': result = getAllTeachersGrouped(); break;

      case 'getDashboardStats': result = getDashboardStats(); break;

      case 'createExam': result = createExam(data); break;
      case 'getExams': result = getExams(); break;
      case 'updateExam': result = updateExam(data); break;

      case 'addRoom': result = addRoom(data); break;
      case 'getRooms': result = getRooms(data); break;
      case 'updateRoom': result = updateRoom(data); break;
      case 'deleteRoom': result = deleteRoom(data); break;

      case 'generateSeatPlan': result = generateSeatPlan(data); break;
      case 'getSeatPlan': result = getSeatPlan(data); break;

      case 'autoAssignInvigilators': result = autoAssignInvigilators(data); break;
      case 'getInvigilators': result = getInvigilators(data); break;

      case 'getStudentFeeReport': result = getStudentFeeReport(data); break;
      case 'submitPayment': result = submitPayment(data); break;
      case 'updatePaymentStatus': result = updatePaymentStatus(data); break;

      default:
        result = { success: false, message: 'অজানা action' };
    }
  } finally {
    lock.releaseLock();
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================================
// অথেন্টিকেশন
// Schools কলাম: school_id(0) school_name(1) email(2) password_hash(3) phone(4) status(5) created_date(6) admin_note(7)
// Users কলাম: user_id(0) name(1) email(2) password_hash(3) role(4) school_id(5) created_date(6) status(7)
// =====================================================================

function registerSchool(data) {
  if (!data.phone) return { success: false, message: 'মোবাইল নম্বর আবশ্যক' };

  const sheet = getSheet('Schools');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === data.email) {
      return { success: false, message: 'এই ইমেইল দিয়ে আগে থেকে রেজিস্টার করা আছে' };
    }
  }

  const schoolId = generateId('SCH');
  sheet.appendRow([schoolId, data.schoolName, data.email, hashPassword(data.password), data.phone, 'pending', new Date(), '']);
  return { success: true, message: 'রেজিস্ট্রেশন সফল, Admin অনুমোদনের অপেক্ষায় আছে', schoolId };
}

function login(data) {
  const hashedInput = hashPassword(data.password);

  const userRows = getSheet('Users').getDataRange().getValues();
  for (let i = 1; i < userRows.length; i++) {
    if (userRows[i][2] === data.email && userRows[i][3] === hashedInput) {
      const status = userRows[i][7]; // blank/undefined ধরা হবে active হিসেবে (পুরনো রেকর্ড ব্যাকওয়ার্ড কম্প্যাটিবিলিটি)
      if (status === 'disabled') {
        return { success: false, message: 'আপনার অ্যাকাউন্ট নিষ্ক্রিয় করা হয়েছে। Super Admin-এর সাথে যোগাযোগ করুন।' };
      }
      return {
        success: true, role: userRows[i][4], schoolId: userRows[i][5] || null,
        schoolName: userRows[i][1] // আসল নাম (Admin/Super Admin/Teacher)
      };
    }
  }

  const schoolRows = getSheet('Schools').getDataRange().getValues();
  for (let i = 1; i < schoolRows.length; i++) {
    if (schoolRows[i][2] === data.email) {
      if (schoolRows[i][3] !== hashedInput) continue;
      const status = schoolRows[i][5];
      const note = schoolRows[i][7];
      if (status === 'pending') return { success: false, message: 'আপনার স্কুল এখনো Admin অনুমোদনের অপেক্ষায় আছে।' };
      if (status === 'rejected') return { success: false, message: 'আপনার আবেদনটি বাতিল করা হয়েছে।' + (note ? ' কারণ: ' + note : '') };
      if (status === 'returned') return { success: false, message: 'আপনার আবেদনটি সংশোধনের জন্য ফেরত পাঠানো হয়েছে।' + (note ? ' নোট: ' + note : '') + ' অনুগ্রহ করে Admin-এর সাথে যোগাযোগ করুন।' };
      if (status !== 'approved') return { success: false, message: 'আপনার স্কুল এখনো সক্রিয় নয়।' };
      return { success: true, role: 'school_admin', schoolId: schoolRows[i][0], schoolName: schoolRows[i][1] };
    }
  }
  return { success: false, message: 'ইমেইল বা পাসওয়ার্ড ভুল' };
}

// =====================================================================
// স্কুল — Super Admin: view all, edit, delete | Admin: approve/reject/return
// =====================================================================

function getAllSchools() {
  const rows = getSheet('Schools').getDataRange().getValues();
  const schools = [];
  for (let i = 1; i < rows.length; i++) {
    schools.push({
      schoolId: rows[i][0], schoolName: rows[i][1], email: rows[i][2],
      phone: rows[i][4], status: rows[i][5], createdDate: rows[i][6], adminNote: rows[i][7]
    });
  }
  return { success: true, schools };
}

// একটা স্কুল ক্লিক করলে তার সব ডেটা (স্টুডেন্ট + টিচার) — সুপার অ্যাডমিনের জন্য
function getSchoolDetail(data) {
  const rows = getSheet('Schools').getDataRange().getValues();
  let school = null;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.schoolId) {
      school = { schoolId: rows[i][0], schoolName: rows[i][1], email: rows[i][2], phone: rows[i][4], status: rows[i][5], adminNote: rows[i][7] };
      break;
    }
  }
  if (!school) return { success: false, message: 'স্কুল খুঁজে পাওয়া যায়নি' };

  const students = getStudents({ schoolId: data.schoolId }).students;
  const teachers = getTeachers({ schoolId: data.schoolId }).teachers;
  return { success: true, school, students, teachers };
}

// শুধু Admin — অনুমোদন
function approveSchool(data) {
  if (!isAdminOrSuper(data)) return { success: false, message: ADMIN_ONLY_MSG };
  const sheet = getSheet('Schools');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.schoolId) {
      sheet.getRange(i + 1, 6).setValue('approved');
      logAudit(data.requesterRole, data.requesterName, 'approve', rows[i][1], '');
      return { success: true, message: 'স্কুল অ্যাপ্রুভ করা হয়েছে' };
    }
  }
  return { success: false, message: 'স্কুল খুঁজে পাওয়া যায়নি' };
}

// শুধু Admin — বাতিল (কারণসহ)
function rejectSchool(data) {
  if (!isAdminOrSuper(data)) return { success: false, message: ADMIN_ONLY_MSG };
  const sheet = getSheet('Schools');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.schoolId) {
      sheet.getRange(i + 1, 6).setValue('rejected');
      sheet.getRange(i + 1, 8).setValue(data.note || '');
      logAudit(data.requesterRole, data.requesterName, 'reject', rows[i][1], data.note || '');
      return { success: true, message: 'স্কুল রিজেক্ট করা হয়েছে' };
    }
  }
  return { success: false, message: 'স্কুল খুঁজে পাওয়া যায়নি' };
}

// শুধু Admin — সংশোধনের জন্য ফেরত পাঠানো
function returnForCorrection(data) {
  if (!isAdminOrSuper(data)) return { success: false, message: ADMIN_ONLY_MSG };
  const sheet = getSheet('Schools');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.schoolId) {
      sheet.getRange(i + 1, 6).setValue('returned');
      sheet.getRange(i + 1, 8).setValue(data.note || '');
      logAudit(data.requesterRole, data.requesterName, 'return', rows[i][1], data.note || '');
      return { success: true, message: 'সংশোধনের জন্য ফেরত পাঠানো হয়েছে' };
    }
  }
  return { success: false, message: 'স্কুল খুঁজে পাওয়া যায়নি' };
}

// শুধু Super Admin — স্কুলের তথ্য এডিট
function updateSchool(data) {
  if (!isSuperAdmin(data)) return { success: false, restricted: true, message: RESTRICTED_MSG };
  const sheet = getSheet('Schools');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.schoolId) {
      sheet.getRange(i + 1, 2).setValue(data.schoolName);
      sheet.getRange(i + 1, 3).setValue(data.email);
      sheet.getRange(i + 1, 5).setValue(data.phone);
      return { success: true };
    }
  }
  return { success: false, message: 'স্কুল খুঁজে পাওয়া যায়নি' };
}

// শুধু Super Admin — স্কুল ডিলিট
function deleteSchool(data) {
  if (!isSuperAdmin(data)) return { success: false, restricted: true, message: RESTRICTED_MSG };
  const sheet = getSheet('Schools');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.schoolId) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, message: 'স্কুল খুঁজে পাওয়া যায়নি' };
}

// শুধু Super Admin — কোনো স্কুলের জন্য নতুন পাসওয়ার্ড সেট করা (পুরনোটা কখনো দেখা/ফেরত আনা যায় না, hash থেকে সম্ভব না —
// তাই "দেখা" এর বদলে "রিসেট" করাই একমাত্র নিরাপদ সমাধান)
function resetSchoolPassword(data) {
  if (!isSuperAdmin(data)) return { success: false, message: SUPER_ADMIN_ONLY_MSG };
  if (!data.newPassword || data.newPassword.length < 6) return { success: false, message: 'পাসওয়ার্ড কমপক্ষে ৬ ক্যারেক্টার হতে হবে' };

  const sheet = getSheet('Schools');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.schoolId) {
      sheet.getRange(i + 1, 4).setValue(hashPassword(data.newPassword));
      logAudit('super_admin', data.requesterName, 'reset_password', rows[i][1], '');
      return { success: true };
    }
  }
  return { success: false, message: 'স্কুল খুঁজে পাওয়া যায়নি' };
}

// =====================================================================
// Admin অ্যাকাউন্ট ব্যবস্থাপনা — শুধু Super Admin তৈরি/নিষ্ক্রিয় করতে পারবে
// =====================================================================

function addAdmin(data) {
  if (!isSuperAdmin(data)) return { success: false, message: SUPER_ADMIN_ONLY_MSG };
  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === data.email) return { success: false, message: 'এই ইমেইল দিয়ে আগে থেকে অ্যাকাউন্ট আছে' };
  }
  const userId = generateId('ADM');
  sheet.appendRow([userId, data.name, data.email, hashPassword(data.password), 'admin', '', new Date(), 'active']);
  return { success: true, userId };
}

function getAdmins() {
  const rows = getSheet('Users').getDataRange().getValues();
  const admins = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][4] === 'admin') {
      admins.push({ userId: rows[i][0], name: rows[i][1], email: rows[i][2], status: rows[i][7] || 'active', createdDate: rows[i][6] });
    }
  }
  return { success: true, admins };
}

function updateAdminStatus(data) {
  if (!isSuperAdmin(data)) return { success: false, message: SUPER_ADMIN_ONLY_MSG };
  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.userId) { sheet.getRange(i + 1, 8).setValue(data.status); return { success: true }; }
  }
  return { success: false, message: 'অ্যাডমিন খুঁজে পাওয়া যায়নি' };
}

// শুধু Super Admin — কোনো Admin এর জন্য নতুন পাসওয়ার্ড সেট করা
function resetAdminPassword(data) {
  if (!isSuperAdmin(data)) return { success: false, message: SUPER_ADMIN_ONLY_MSG };
  if (!data.newPassword || data.newPassword.length < 6) return { success: false, message: 'পাসওয়ার্ড কমপক্ষে ৬ ক্যারেক্টার হতে হবে' };

  const sheet = getSheet('Users');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.userId) {
      sheet.getRange(i + 1, 4).setValue(hashPassword(data.newPassword));
      return { success: true };
    }
  }
  return { success: false, message: 'অ্যাডমিন খুঁজে পাওয়া যায়নি' };
}

// =====================================================================
// Audit Log — প্রতিটা Approve/Reject/Return এখানে রেকর্ড হয়
// =====================================================================

function logAudit(actorRole, actorName, action, targetSchoolName, note) {
  getSheet('AuditLog').appendRow([generateId('LOG'), new Date(), actorRole, actorName || '', action, targetSchoolName || '', note || '']);
}

function getAuditLog(data) {
  const rows = getSheet('AuditLog').getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < rows.length; i++) {
    const row = { timestamp: rows[i][1], actorRole: rows[i][2], actorName: rows[i][3], action: rows[i][4], schoolName: rows[i][5], note: rows[i][6] };
    // Admin শুধু নিজের কার্যক্রম দেখবে, Super Admin সব দেখবে
    if (data.requesterRole === 'admin') { if (row.actorName === data.requesterName) logs.push(row); }
    else logs.push(row);
  }
  logs.reverse(); // সাম্প্রতিকটা আগে
  return { success: true, logs };
}

// =====================================================================
// শিক্ষার্থী নিবন্ধন ON/OFF — শুধু Super Admin নিয়ন্ত্রণ করবে
// =====================================================================

// Google Sheets 'true'/'false' স্ট্রিং লিখলেও অনেক সময় নিজে থেকেই আসল Boolean এ রূপান্তর করে ফেলে
// (যেমন তারিখ auto-detect করে), তাই strict string comparison নির্ভরযোগ্য না — এই হেল্পার
// Boolean false, স্ট্রিং 'false'/'FALSE'/'False' — সবকটা ক্ষেত্র নিরাপদে ধরবে
function isFalsyValue(value) {
  return value === false || value === 'false' || value === 'FALSE' || value === 'False';
}

function getStudentRegistrationStatus() {
  const rows = getSheet('Config').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'student_registration_enabled') return { success: true, enabled: !isFalsyValue(rows[i][1]) };
  }
  return { success: true, enabled: true }; // কনফিগ না থাকলে ডিফল্ট ON
}

function setStudentRegistrationStatus(data) {
  if (!isSuperAdmin(data)) return { success: false, message: SUPER_ADMIN_ONLY_MSG };
  const sheet = getSheet('Config');
  const rows = sheet.getDataRange().getValues();
  const valueToStore = data.enabled === true || data.enabled === 'true' ? true : false; // আসল Boolean হিসেবেই সেভ করা হচ্ছে
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'student_registration_enabled') {
      sheet.getRange(i + 1, 2).setValue(valueToStore);
      return { success: true };
    }
  }
  sheet.appendRow(['student_registration_enabled', valueToStore]);
  return { success: true };
}

// =====================================================================
// ক্লাস — শুধু Super Admin তৈরি করবে (গ্লোবাল লিস্ট)
// =====================================================================

function addClass(data) {
  const sheet = getSheet('Classes');
  const classId = generateId('CLS');
  sheet.appendRow([classId, data.className]);
  return { success: true, classId };
}

function getClasses() {
  const rows = getSheet('Classes').getDataRange().getValues();
  const classes = [];
  for (let i = 1; i < rows.length; i++) classes.push({ classId: rows[i][0], className: rows[i][1] });
  return { success: true, classes };
}

// =====================================================================
// স্টুডেন্ট
// Students কলাম: student_id(0) name(1) class_id(2) school_id(3) roll_no(4) contact(5) father_name(6) mother_name(7)
// =====================================================================

function addStudent(data) {
  const regStatus = getStudentRegistrationStatus();
  if (!regStatus.enabled) return { success: false, message: 'শিক্ষার্থী নিবন্ধন বর্তমানে বন্ধ আছে। Super Admin চালু করলে আবার যোগ করতে পারবেন।' };

  const sheet = getSheet('Students');
  const studentId = generateId('STU');
  sheet.appendRow([studentId, data.name, data.classId, data.schoolId, '', data.contact || '', data.fatherName || '', data.motherName || '']);
  return { success: true, studentId };
}

function getStudents(data) {
  const studentRows = getSheet('Students').getDataRange().getValues();
  const classRows = getSheet('Classes').getDataRange().getValues();
  const classMap = {};
  for (let i = 1; i < classRows.length; i++) classMap[classRows[i][0]] = classRows[i][1];

  const students = [];
  for (let i = 1; i < studentRows.length; i++) {
    if (studentRows[i][3] === data.schoolId) {
      students.push({
        studentId: studentRows[i][0], name: studentRows[i][1], classId: studentRows[i][2],
        className: classMap[studentRows[i][2]] || '—', rollNo: studentRows[i][4], contact: studentRows[i][5],
        fatherName: studentRows[i][6], motherName: studentRows[i][7]
      });
    }
  }
  return { success: true, students };
}

// School Admin (নিজের স্কুলের মধ্যে) বা Super Admin — এডিট করতে পারবে (Delete না)
function updateStudent(data) {
  const sheet = getSheet('Students');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.studentId) {
      if (data.requesterRole === 'school_admin' && rows[i][3] !== data.requesterSchoolId) {
        return { success: false, message: 'অনুমতি নেই' };
      }
      sheet.getRange(i + 1, 2).setValue(data.name);
      sheet.getRange(i + 1, 3).setValue(data.classId);
      sheet.getRange(i + 1, 6).setValue(data.contact || '');
      sheet.getRange(i + 1, 7).setValue(data.fatherName || '');
      sheet.getRange(i + 1, 8).setValue(data.motherName || '');
      return { success: true };
    }
  }
  return { success: false, message: 'স্টুডেন্ট খুঁজে পাওয়া যায়নি' };
}

// শুধু Super Admin ডিলিট করতে পারবে
function deleteStudent(data) {
  if (!isSuperAdmin(data)) return { success: false, restricted: true, message: RESTRICTED_MSG };
  const sheet = getSheet('Students');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.studentId) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, message: 'স্টুডেন্ট খুঁজে পাওয়া যায়নি' };
}

// শুধু Super Admin — রোল নম্বর বসানো (বাংলা সংখ্যায়)
function updateStudentRoll(data) {
  const sheet = getSheet('Students');
  const rows = sheet.getDataRange().getValues();
  const banglaRoll = toBanglaDigits(data.rollNo);
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.studentId) { sheet.getRange(i + 1, 5).setValue(banglaRoll); return { success: true, rollNo: banglaRoll }; }
  }
  return { success: false, message: 'স্টুডেন্ট খুঁজে পাওয়া যায়নি' };
}

function getAllStudentsGrouped() {
  const studentRows = getSheet('Students').getDataRange().getValues();
  const classRows = getSheet('Classes').getDataRange().getValues();
  const schoolRows = getSheet('Schools').getDataRange().getValues();

  const classMap = {}; for (let i = 1; i < classRows.length; i++) classMap[classRows[i][0]] = classRows[i][1];
  const schoolMap = {}; for (let i = 1; i < schoolRows.length; i++) schoolMap[schoolRows[i][0]] = schoolRows[i][1];

  const students = [];
  for (let i = 1; i < studentRows.length; i++) {
    students.push({
      studentId: studentRows[i][0], name: studentRows[i][1], schoolId: studentRows[i][3], classId: studentRows[i][2],
      schoolName: schoolMap[studentRows[i][3]] || '—', className: classMap[studentRows[i][2]] || '—',
      rollNo: studentRows[i][4], contact: studentRows[i][5], fatherName: studentRows[i][6], motherName: studentRows[i][7]
    });
  }
  return { success: true, students };
}

// রোল নম্বর অ্যাসাইনমেন্টের জন্য — Class অনুযায়ী সাজানো (Class 1 এর সব স্কুলের ছাত্র, তারপর Class 2...)
function getStudentsForRollAssignment() {
  const studentRows = getSheet('Students').getDataRange().getValues();
  const classRows = getSheet('Classes').getDataRange().getValues();
  const schoolRows = getSheet('Schools').getDataRange().getValues();

  const schoolMap = {}; for (let i = 1; i < schoolRows.length; i++) schoolMap[schoolRows[i][0]] = schoolRows[i][1];

  const byClass = {};
  for (let i = 1; i < classRows.length; i++) {
    byClass[classRows[i][0]] = { classId: classRows[i][0], className: classRows[i][1], students: [] };
  }

  const orphan = { classId: '', className: 'অনির্ধারিত ক্লাস', students: [] };

  for (let i = 1; i < studentRows.length; i++) {
    const classId = studentRows[i][2];
    const entry = {
      studentId: studentRows[i][0], name: studentRows[i][1],
      schoolName: schoolMap[studentRows[i][3]] || '—', rollNo: studentRows[i][4]
    };
    if (byClass[classId]) byClass[classId].students.push(entry);
    else orphan.students.push(entry);
  }

  // Classes ট্যাবে যে ক্রমে ক্লাস তৈরি করা হয়েছে সেই ক্রম অনুযায়ী (Class 1, Class 2...)
  const orderedClassIds = [];
  for (let i = 1; i < classRows.length; i++) orderedClassIds.push(classRows[i][0]);

  const classes = orderedClassIds.map(cid => byClass[cid]).filter(c => c.students.length > 0);
  if (orphan.students.length > 0) classes.push(orphan);

  return { success: true, classes };
}

// =====================================================================
// টিচার
// =====================================================================

function addTeacher(data) {
  const sheet = getSheet('Teachers');
  const teacherId = generateId('TCH');
  sheet.appendRow([teacherId, data.name, data.subject, data.classId, data.schoolId, data.contact]);
  return { success: true, teacherId };
}

function getTeachers(data) {
  const teacherRows = getSheet('Teachers').getDataRange().getValues();
  const classRows = getSheet('Classes').getDataRange().getValues();
  const classMap = {}; for (let i = 1; i < classRows.length; i++) classMap[classRows[i][0]] = classRows[i][1];

  const teachers = [];
  for (let i = 1; i < teacherRows.length; i++) {
    if (teacherRows[i][4] === data.schoolId) {
      teachers.push({
        teacherId: teacherRows[i][0], name: teacherRows[i][1], subject: teacherRows[i][2],
        classId: teacherRows[i][3], className: classMap[teacherRows[i][3]] || '—', contact: teacherRows[i][5]
      });
    }
  }
  return { success: true, teachers };
}

function updateTeacher(data) {
  const sheet = getSheet('Teachers');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.teacherId) {
      if (data.requesterRole === 'school_admin' && rows[i][4] !== data.requesterSchoolId) {
        return { success: false, message: 'অনুমতি নেই' };
      }
      sheet.getRange(i + 1, 2).setValue(data.name);
      sheet.getRange(i + 1, 3).setValue(data.subject);
      sheet.getRange(i + 1, 4).setValue(data.classId);
      sheet.getRange(i + 1, 6).setValue(data.contact || '');
      return { success: true };
    }
  }
  return { success: false, message: 'টিচার খুঁজে পাওয়া যায়নি' };
}

function deleteTeacher(data) {
  if (!isSuperAdmin(data)) return { success: false, restricted: true, message: RESTRICTED_MSG };
  const sheet = getSheet('Teachers');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.teacherId) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, message: 'টিচার খুঁজে পাওয়া যায়নি' };
}

function getAllTeachersGrouped() {
  const teacherRows = getSheet('Teachers').getDataRange().getValues();
  const classRows = getSheet('Classes').getDataRange().getValues();
  const schoolRows = getSheet('Schools').getDataRange().getValues();

  const classMap = {}; for (let i = 1; i < classRows.length; i++) classMap[classRows[i][0]] = classRows[i][1];
  const schoolMap = {}; for (let i = 1; i < schoolRows.length; i++) schoolMap[schoolRows[i][0]] = schoolRows[i][1];

  const teachers = [];
  for (let i = 1; i < teacherRows.length; i++) {
    teachers.push({
      teacherId: teacherRows[i][0], name: teacherRows[i][1], subject: teacherRows[i][2],
      schoolId: teacherRows[i][4], schoolName: schoolMap[teacherRows[i][4]] || '—', classId: teacherRows[i][3],
      className: classMap[teacherRows[i][3]] || '—', contact: teacherRows[i][5]
    });
  }
  return { success: true, teachers };
}

// =====================================================================
// ড্যাশবোর্ড স্ট্যাট
// =====================================================================

function getDashboardStats() {
  const schoolRows = getSheet('Schools').getDataRange().getValues();
  const studentRows = getSheet('Students').getDataRange().getValues();
  const teacherRows = getSheet('Teachers').getDataRange().getValues();

  let approvedSchools = 0;
  for (let i = 1; i < schoolRows.length; i++) if (schoolRows[i][5] === 'approved') approvedSchools++;

  return {
    success: true, totalSchools: approvedSchools,
    totalStudents: Math.max(studentRows.length - 1, 0),
    totalTeachers: Math.max(teacherRows.length - 1, 0)
  };
}

// =====================================================================
// পরীক্ষা — fee_per_student শুধু Super Admin সেট করবে
// Exams কলাম: exam_id(0) exam_name(1) exam_date(2) status(3) fee_per_student(4)
// =====================================================================

function createExam(data) {
  const sheet = getSheet('Exams');
  const examId = generateId('EXM');
  sheet.appendRow([examId, data.examName, data.examDate, 'draft', Number(data.feePerStudent) || 0]);
  return { success: true, examId };
}

function getExams() {
  const rows = getSheet('Exams').getDataRange().getValues();
  const exams = [];
  for (let i = 1; i < rows.length; i++) {
    exams.push({ examId: rows[i][0], examName: rows[i][1], examDate: rows[i][2], status: rows[i][3], feePerStudent: Number(rows[i][4]) || 0 });
  }
  return { success: true, exams };
}

function updateExam(data) {
  const sheet = getSheet('Exams');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.examId) {
      sheet.getRange(i + 1, 2).setValue(data.examName);
      sheet.getRange(i + 1, 3).setValue(data.examDate);
      sheet.getRange(i + 1, 5).setValue(Number(data.feePerStudent) || 0);
      return { success: true };
    }
  }
  return { success: false, message: 'পরীক্ষা খুঁজে পাওয়া যায়নি' };
}

// =====================================================================
// রুম — Add / Edit / Delete
// =====================================================================

function addRoom(data) {
  const sheet = getSheet('Rooms');
  const roomId = generateId('RM');
  sheet.appendRow([roomId, data.examId, data.roomName, data.capacity]);
  return { success: true, roomId };
}

function getRooms(data) {
  const rows = getSheet('Rooms').getDataRange().getValues();
  const rooms = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === data.examId) rooms.push({ roomId: rows[i][0], roomName: rows[i][2], capacity: Number(rows[i][3]) });
  }
  return { success: true, rooms };
}

function updateRoom(data) {
  const sheet = getSheet('Rooms');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.roomId) {
      sheet.getRange(i + 1, 3).setValue(data.roomName);
      sheet.getRange(i + 1, 4).setValue(data.capacity);
      return { success: true };
    }
  }
  return { success: false, message: 'রুম খুঁজে পাওয়া যায়নি' };
}

function deleteRoom(data) {
  const sheet = getSheet('Rooms');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.roomId) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, message: 'রুম খুঁজে পাওয়া যায়নি' };
}

// =====================================================================
// সিট প্ল্যান
// =====================================================================

function generateSeatPlan(data) {
  const examId = data.examId;
  const rooms = getRooms({ examId }).rooms;
  if (rooms.length === 0) return { success: false, message: 'আগে রুম যোগ করুন' };

  const studentRows = getSheet('Students').getDataRange().getValues();
  const schoolRows = getSheet('Schools').getDataRange().getValues();

  const approvedSchools = {};
  for (let i = 1; i < schoolRows.length; i++) if (schoolRows[i][5] === 'approved') approvedSchools[schoolRows[i][0]] = true;

  const bySchool = {};
  for (let i = 1; i < studentRows.length; i++) {
    const schoolId = studentRows[i][3];
    if (approvedSchools[schoolId]) {
      if (!bySchool[schoolId]) bySchool[schoolId] = [];
      bySchool[schoolId].push({ studentId: studentRows[i][0], name: studentRows[i][1], schoolId, rollNo: studentRows[i][4] });
    }
  }

  let totalStudents = 0;
  let schoolChunks = Object.keys(bySchool).map(sid => {
    totalStudents += bySchool[sid].length;
    return { schoolId: sid, students: bySchool[sid].slice() };
  });

  const totalCapacity = rooms.reduce((s, r) => s + r.capacity, 0);
  if (totalStudents > totalCapacity) {
    return { success: false, message: `মোট শিক্ষার্থী (${totalStudents}) রুমের ধারণক্ষমতার (${totalCapacity}) চেয়ে বেশি` };
  }

  const roomState = rooms.map(r => ({ roomId: r.roomId, roomName: r.roomName, capacity: r.capacity, remaining: r.capacity, assigned: [] }));

  // বড় থেকে ছোট স্কুল সাজানো (Best Fit Decreasing হিউরিস্টিক)
  schoolChunks.sort((a, b) => b.students.length - a.students.length);

  // ===== ধাপ ১: প্রতিটা স্কুলকে সম্পূর্ণভাবে এক রুমে বসানোর চেষ্টা — Exact Match > Best Fit =====
  const unplaced = [];
  schoolChunks.forEach(chunk => {
    const size = chunk.students.length;
    const candidates = roomState.filter(r => r.remaining >= size);
    if (candidates.length === 0) { unplaced.push(chunk); return; }

    candidates.sort((a, b) => {
      const leftoverA = a.remaining - size, leftoverB = b.remaining - size;
      if (leftoverA !== leftoverB) return leftoverA - leftoverB;       // কম খালি আসন আগে (Exact Match/Best Fit)
      return a.assigned.length - b.assigned.length;                    // কম স্কুলযুক্ত রুম আগে
    });

    const target = candidates[0];
    target.assigned.push({ schoolId: chunk.schoolId, students: chunk.students });
    target.remaining -= size;
  });

  // ===== ধাপ ২: যেসব স্কুল কোনো একক রুমে সম্পূর্ণ বসেনি — বাকি থাকা capacity-তে বিন-প্যাকিং (একাধিক স্কুল মিলিয়ে) =====
  let pending = unplaced;
  roomState.forEach(room => {
    while (room.remaining > 0 && pending.length > 0) {
      pending.sort((a, b) => b.students.length - a.students.length);
      const fitIndex = pending.findIndex(c => c.students.length > 0 && c.students.length <= room.remaining);
      if (fitIndex !== -1) {
        const chunk = pending[fitIndex];
        room.assigned.push({ schoolId: chunk.schoolId, students: chunk.students });
        room.remaining -= chunk.students.length;
        pending.splice(fitIndex, 1);
      } else {
        const chunk = pending[pending.length - 1]; // সবচেয়ে ছোট স্কুল ভাঙা হবে (শেষ অবলম্বন)
        const taken = chunk.students.splice(0, room.remaining);
        room.assigned.push({ schoolId: chunk.schoolId, students: taken });
        room.remaining = 0;
        if (chunk.students.length === 0) pending.pop();
      }
    }
  });

  // ===== ধাপ ৩: Final Optimization — একই স্কুল একাধিক রুমে ভাগ হয়ে থাকলে সমান-সাইজ swap দিয়ে একত্র করা (ধারণক্ষমতা কখনো ভাঙে না) =====
  consolidateSplitSchools(roomState);

  // ===== SeatPlan শীটে লেখা =====
  const seatSheet = getSheet('SeatPlan');
  const existingRows = seatSheet.getDataRange().getValues();
  for (let i = existingRows.length - 1; i >= 1; i--) if (existingRows[i][1] === examId) seatSheet.deleteRow(i + 1);

  const roomSummary = [];
  let totalSeated = 0;
  roomState.forEach(room => {
    let seatedHere = 0;
    room.assigned.forEach(group => {
      group.students.forEach(s => {
        seatSheet.appendRow([generateId('SEAT'), examId, room.roomId, s.studentId, s.name, s.schoolId, s.rollNo]);
        seatedHere++;
      });
    });
    totalSeated += seatedHere;
    roomSummary.push({ roomName: room.roomName, seated: seatedHere, capacity: room.capacity });
  });

  const examSheet = getSheet('Exams');
  const examRows = examSheet.getDataRange().getValues();
  for (let i = 1; i < examRows.length; i++) if (examRows[i][0] === examId) examSheet.getRange(i + 1, 4).setValue('seated');

  return { success: true, totalStudents: totalSeated, roomSummary };
}

// একই স্কুলের শিক্ষার্থী একাধিক রুমে ছড়িয়ে থাকলে, সমান আকারের অন্য একটা স্কুলের গ্রুপের সাথে
// অদল-বদল (swap) করে যতটা সম্ভব একত্র করার চেষ্টা — সমান সাইজ swap বলে room capacity কখনো ভাঙে না
function consolidateSplitSchools(roomState) {
  let changed = true;
  let iterations = 0;
  while (changed && iterations < 20) {
    changed = false;
    iterations++;

    const locations = {};
    roomState.forEach((room, ri) => {
      room.assigned.forEach((group, gi) => {
        if (!locations[group.schoolId]) locations[group.schoolId] = [];
        locations[group.schoolId].push({ roomIdx: ri, groupIdx: gi, size: group.students.length });
      });
    });

    for (const schoolId in locations) {
      const locs = locations[schoolId];
      if (locs.length < 2) continue;

      locs.sort((a, b) => b.size - a.size);
      const home = locs[0];

      for (let k = 1; k < locs.length; k++) {
        const away = locs[k];
        if (away.roomIdx === home.roomIdx) continue;

        const homeRoom = roomState[home.roomIdx];
        const awayRoom = roomState[away.roomIdx];

        const swapIdx = homeRoom.assigned.findIndex(g => g.schoolId !== schoolId && g.students.length === away.size);
        if (swapIdx !== -1) {
          const temp = homeRoom.assigned[swapIdx];
          homeRoom.assigned[swapIdx] = awayRoom.assigned[away.groupIdx];
          awayRoom.assigned[away.groupIdx] = temp;
          changed = true;
          break;
        }
      }
      if (changed) break;
    }
  }
}

// বাংলা সংখ্যাকে ইংরেজিতে (শুধু sorting এর জন্য)
function banglaToEnglishDigits(input) {
  const map = { '০':'0','১':'1','২':'2','৩':'3','৪':'4','৫':'5','৬':'6','৭':'7','৮':'8','৯':'9' };
  return String(input || '').replace(/[০-৯]/g, d => map[d]);
}
function rollSortValue(roll) {
  const num = parseInt(banglaToEnglishDigits(roll), 10);
  return isNaN(num) ? Infinity : num;
}

// সিট প্ল্যান দেখা — প্রতি রুমে প্রতিটা স্কুলের অধীনে রোল নম্বর, মোট শিক্ষার্থী, খালি আসন
function getSeatPlan(data) {
  const rows = getSheet('SeatPlan').getDataRange().getValues();
  const roomRows = getSheet('Rooms').getDataRange().getValues();
  const schoolRows = getSheet('Schools').getDataRange().getValues();
  const studentRows = getSheet('Students').getDataRange().getValues();

  const roomMap = {};
  for (let i = 1; i < roomRows.length; i++) {
    if (roomRows[i][1] === data.examId) roomMap[roomRows[i][0]] = { roomName: roomRows[i][2], capacity: Number(roomRows[i][3]) };
  }
  const schoolNameMap = {};
  for (let i = 1; i < schoolRows.length; i++) schoolNameMap[schoolRows[i][0]] = schoolRows[i][1];

  // Live roll number lookup — রোল যদি সিট প্ল্যান তৈরির পরে বসানো হয়, তাও সঠিক দেখাবে
  const liveRollMap = {};
  for (let i = 1; i < studentRows.length; i++) liveRollMap[studentRows[i][0]] = studentRows[i][4];

  const byRoom = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] !== data.examId) continue;
    const roomId = rows[i][2], studentId = rows[i][3], schoolId = rows[i][5];
    const roomInfo = roomMap[roomId] || { roomName: roomId, capacity: 0 };

    if (!byRoom[roomId]) byRoom[roomId] = { roomName: roomInfo.roomName, capacity: roomInfo.capacity, bySchool: {}, total: 0 };
    if (!byRoom[roomId].bySchool[schoolId]) byRoom[roomId].bySchool[schoolId] = { schoolName: schoolNameMap[schoolId] || '—', rolls: [] };

    const roll = liveRollMap.hasOwnProperty(studentId) ? liveRollMap[studentId] : rows[i][6];
    byRoom[roomId].bySchool[schoolId].rolls.push(roll || '—');
    byRoom[roomId].total++;
  }

  const roomsOut = Object.values(byRoom).map(r => {
    const schools = Object.values(r.bySchool).map(s => {
      s.rolls.sort((a, b) => rollSortValue(a) - rollSortValue(b));
      return s;
    });
    return {
      roomName: r.roomName, capacity: r.capacity, schools,
      totalStudents: r.total, emptySeats: Math.max(r.capacity - r.total, 0)
    };
  });

  return { success: true, rooms: roomsOut };
}

// =====================================================================
// ইনভিজিলেটর — প্রতি রুমে কমপক্ষে ২ জন, ভিন্ন স্কুলের
// =====================================================================

function autoAssignInvigilators(data) {
  const examId = data.examId;
  const seatRows = getSheet('SeatPlan').getDataRange().getValues();
  const roomRows = getSheet('Rooms').getDataRange().getValues();

  const roomSchoolMap = {}, roomNameMap = {};
  for (let i = 1; i < roomRows.length; i++) {
    if (roomRows[i][1] === examId) { roomSchoolMap[roomRows[i][0]] = new Set(); roomNameMap[roomRows[i][0]] = roomRows[i][2]; }
  }
  for (let i = 1; i < seatRows.length; i++) {
    if (seatRows[i][1] === examId) { const roomId = seatRows[i][2]; if (roomSchoolMap[roomId]) roomSchoolMap[roomId].add(seatRows[i][5]); }
  }

  const teacherRows = getSheet('Teachers').getDataRange().getValues();
  let teacherPool = [];
  for (let i = 1; i < teacherRows.length; i++) {
    teacherPool.push({ teacherId: teacherRows[i][0], name: teacherRows[i][1], schoolId: teacherRows[i][4], phone: teacherRows[i][5] });
  }
  teacherPool = shuffleArray(teacherPool);

  const invigilatorSheet = getSheet('Invigilators');
  const existingRows = invigilatorSheet.getDataRange().getValues();
  for (let i = existingRows.length - 1; i >= 1; i--) if (existingRows[i][1] === examId) invigilatorSheet.deleteRow(i + 1);

  const MIN_PER_ROOM = 2;
  const usedTeacherIds = new Set();
  const roomResults = [], shortRooms = [];

  Object.keys(roomSchoolMap).forEach(roomId => {
    const excludedSchools = roomSchoolMap[roomId];
    const assignedForRoom = [];

    for (let n = 0; n < MIN_PER_ROOM; n++) {
      const eligible = teacherPool.find(t => !excludedSchools.has(t.schoolId) && !usedTeacherIds.has(t.teacherId));
      if (eligible) {
        usedTeacherIds.add(eligible.teacherId);
        assignedForRoom.push(eligible);
        invigilatorSheet.appendRow([generateId('INV'), examId, roomId, eligible.teacherId, eligible.name, eligible.schoolId, eligible.phone]);
      }
    }

    roomResults.push({ roomName: roomNameMap[roomId], invigilators: assignedForRoom.map(t => ({ name: t.name, phone: t.phone })) });
    if (assignedForRoom.length < MIN_PER_ROOM) shortRooms.push(`${roomNameMap[roomId]} (${assignedForRoom.length}/${MIN_PER_ROOM} জন পাওয়া গেছে)`);
  });

  return { success: true, roomResults, shortRooms };
}

function getInvigilators(data) {
  const rows = getSheet('Invigilators').getDataRange().getValues();
  const roomRows = getSheet('Rooms').getDataRange().getValues();
  const schoolRows = getSheet('Schools').getDataRange().getValues();

  const roomMap = {}; for (let i = 1; i < roomRows.length; i++) roomMap[roomRows[i][0]] = roomRows[i][2];
  const schoolMap = {}; for (let i = 1; i < schoolRows.length; i++) schoolMap[schoolRows[i][0]] = schoolRows[i][1];

  const byRoom = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === data.examId) {
      const roomName = roomMap[rows[i][2]] || rows[i][2];
      if (!byRoom[roomName]) byRoom[roomName] = [];
      byRoom[roomName].push({ name: rows[i][4], phone: rows[i][6], schoolName: schoolMap[rows[i][5]] || '—' });
    }
  }
  return { success: true, list: Object.keys(byRoom).map(roomName => ({ roomName, invigilators: byRoom[roomName] })) };
}

// =====================================================================
// ফি রিপোর্ট ও পেমেন্ট
// Payments কলাম: payment_id(0) exam_id(1) school_id(2) amount(3) method(4) status(5) payment_date(6) note(7)
// =====================================================================

function getStudentFeeReport(data) {
  const examRows = getSheet('Exams').getDataRange().getValues();
  let feePerStudent = 0, examName = '';
  for (let i = 1; i < examRows.length; i++) {
    if (examRows[i][0] === data.examId) { feePerStudent = Number(examRows[i][4]) || 0; examName = examRows[i][1]; }
  }

  const schoolRows = getSheet('Schools').getDataRange().getValues();
  const studentRows = getSheet('Students').getDataRange().getValues();
  const paymentRows = getSheet('Payments').getDataRange().getValues();

  const countMap = {};
  for (let i = 1; i < studentRows.length; i++) { const sid = studentRows[i][3]; countMap[sid] = (countMap[sid] || 0) + 1; }

  const paymentMap = {};
  for (let i = 1; i < paymentRows.length; i++) {
    if (paymentRows[i][1] === data.examId) {
      paymentMap[paymentRows[i][2]] = { amount: paymentRows[i][3], method: paymentRows[i][4], status: paymentRows[i][5], date: paymentRows[i][6] };
    }
  }

  const report = [];
  for (let i = 1; i < schoolRows.length; i++) {
    if (schoolRows[i][5] !== 'approved') continue;
    const schoolId = schoolRows[i][0];
    const studentCount = countMap[schoolId] || 0;
    const payment = paymentMap[schoolId] || null;
    report.push({
      schoolId, schoolName: schoolRows[i][1], phone: schoolRows[i][4],
      studentCount, totalFee: studentCount * feePerStudent,
      paymentStatus: payment ? payment.status : 'pending',
      paymentMethod: payment ? payment.method : '', amountPaid: payment ? payment.amount : 0
    });
  }
  return { success: true, examName, feePerStudent, report };
}

function submitPayment(data) {
  const sheet = getSheet('Payments');
  const paymentId = generateId('PAY');
  sheet.appendRow([paymentId, data.examId, data.schoolId, data.amount, data.method, 'pending', new Date(), data.note || '']);
  return { success: true, paymentId };
}

function updatePaymentStatus(data) {
  if (!isSuperAdmin(data)) return { success: false, restricted: true, message: RESTRICTED_MSG };
  const sheet = getSheet('Payments');
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] === data.examId && rows[i][2] === data.schoolId) {
      sheet.getRange(i + 1, 6).setValue(data.status);
      return { success: true };
    }
  }
  return { success: false, message: 'পেমেন্ট রেকর্ড পাওয়া যায়নি' };
}