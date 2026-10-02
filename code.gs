function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("NPRU Emergency Alert System")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: เชื่อมต่อและเตรียมโครงสร้าง Google Sheet อัตโนมัติ
function getDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userSheet = ss.getSheetByName("Users");
  let incidentSheet = ss.getSheetByName("Incidents");

  if (!userSheet) {
    userSheet = ss.insertSheet("Users");
    userSheet.appendRow([
      "Email",
      "Password",
      "Name",
      "ID",
      "Group",
      "Faculty",
      "Department",
      "Role",
      "Status",
      "PdpaConsent",
      "CreatedAt",
    ]);
  }
  if (!incidentSheet) {
    incidentSheet = ss.insertSheet("Incidents");
    incidentSheet.appendRow([
      "IncidentID",
      "Type",
      "Location",
      "Description",
      "ReporterEmail",
      "Status",
      "AttachmentUrl",
      "CreatedAt",
    ]);
  }

  return { ss, userSheet, incidentSheet };
}

// Helper: ตรวจสอบสิทธิ์ Admin
function checkIsAdmin(userEmail) {
  if (!userEmail) return false;
  try {
    const { userSheet } = getDb();
    const data = userSheet.getDataRange().getValues();
    const emailLower = userEmail.toString().trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === emailLower) {
        return data[i][7] === "Admin";
      }
    }
  } catch (err) {
    return false;
  }
  return false;
}

// ----------------- USER AUTH & PROFILE -----------------
function loginUser(email, password) {
  try {
    const { userSheet } = getDb();
    const data = userSheet.getDataRange().getValues();
    const emailLower = email.toString().trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0].toString().trim().toLowerCase() === emailLower) {
        if (row[8] === "Inactive") {
          return {
            success: false,
            message: "บัญชีนี้ถูกระงับการใช้งาน กรุณาติดต่อ Admin",
          };
        }
        if (row[1].toString() === password.toString()) {
          return {
            success: true,
            data: {
              Email: row[0].toString().trim(),
              Name: row[2].toString(),
              ID: row[3].toString(),
              Group: row[4] ? row[4].toString() : "",
              Faculty: row[5] ? row[5].toString() : "",
              Department: row[6] ? row[6].toString() : "",
              Role: row[7] ? row[7].toString() : "User",
              Status: row[8] ? row[8].toString() : "Active",
            },
          };
        } else {
          return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
        }
      }
    }
    return {
      success: false,
      isNew: true,
      message: "ไม่พบบัญชีผู้ใช้นี้ในระบบ",
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function registerUser(userData) {
  try {
    const { userSheet } = getDb();
    const data = userSheet.getDataRange().getValues();
    const emailLower = userData.email.toString().trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === emailLower) {
        return { success: false, message: "อีเมลนี้ถูกลงทะเบียนไว้ในระบบแล้ว" };
      }
    }

    const role = data.length === 1 ? "Admin" : "User";
    const status = "Active";
    const createdAt = Utilities.formatDate(
      new Date(),
      "Asia/Bangkok",
      "yyyy-MM-dd HH:mm:ss",
    );

    userSheet.appendRow([
      emailLower,
      userData.password.toString(),
      userData.name.toString(),
      userData.id.toString(),
      userData.group ? userData.group.toString() : "",
      userData.faculty.toString(),
      userData.department ? userData.department.toString() : "",
      role,
      status,
      userData.pdpaConsent ? "YES" : "NO",
      createdAt,
    ]);

    return {
      success: true,
      data: {
        Email: emailLower,
        Name: userData.name,
        ID: userData.id,
        Group: userData.group || "",
        Faculty: userData.faculty,
        Department: userData.department || "",
        Role: role,
        Status: status,
      },
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function resetPassword(email, idNumber, newPassword) {
  try {
    const { userSheet } = getDb();
    const data = userSheet.getDataRange().getValues();
    const emailLower = email.toString().trim().toLowerCase();
    const idTrimmed = idNumber.toString().trim();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === emailLower) {
        if (data[i][3].toString().trim() === idTrimmed) {
          userSheet.getRange(i + 1, 2).setValue(newPassword.toString());
          return {
            success: true,
            message: "เปลี่ยนรหัสผ่านใหม่สำเร็จเรียบร้อยแล้ว",
          };
        } else {
          return {
            success: false,
            message: "รหัสนักศึกษา / รหัสบุคลากร ไม่ถูกต้อง",
          };
        }
      }
    }
    return { success: false, message: "ไม่พบอีเมลนี้ในระบบ" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateUserProfile(userData) {
  try {
    const { userSheet } = getDb();
    const data = userSheet.getDataRange().getValues();
    const emailLower = userData.email.toString().trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === emailLower) {
        userSheet.getRange(i + 1, 3).setValue(userData.name);
        userSheet.getRange(i + 1, 4).setValue(userData.id);
        userSheet.getRange(i + 1, 5).setValue(userData.group || "");
        userSheet.getRange(i + 1, 6).setValue(userData.faculty);
        userSheet.getRange(i + 1, 7).setValue(userData.department || "");
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบผู้ใช้ที่ต้องการอัปเดต" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ----------------- INCIDENT MANAGEMENT -----------------
function reportIncident(incData) {
  try {
    const { incidentSheet, userSheet } = getDb();
    const incidentId = "INC-" + Math.floor(100000 + Math.random() * 900000);
    const createdAt = Utilities.formatDate(
      new Date(),
      "Asia/Bangkok",
      "yyyy-MM-dd HH:mm:ss",
    );
    let attachmentUrl = "";

    // ค้นหาชื่อและข้อมูลผู้แจ้งเหตุจากฐานข้อมูล Users
    let reporterDetail = incData.reporterEmail;
    try {
      const users = userSheet.getDataRange().getValues();
      const emailLower = (incData.reporterEmail || "")
        .toString()
        .trim()
        .toLowerCase();
      for (let i = 1; i < users.length; i++) {
        if (users[i][0].toString().trim().toLowerCase() === emailLower) {
          const facultyStr = users[i][5] || "ไม่ระบุคณะ";
          const deptStr = users[i][6] || "ไม่ระบุสาขา";
          reporterDetail = `${users[i][2]} (${facultyStr} / ${deptStr}) }`;
          break;
        }
      }
    } catch (e) {
      // ใช้เป็นอีเมลปกติหากเกิดข้อผิดพลาดในการดึงชื่อ
    }

    if (incData.fileObj) {
      const folderName = "NPRU_Emergency_Uploads";
      let folder;
      const folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder(folderName);
      }

      const base64Data = incData.fileObj.base64.split(",")[1];
      const blob = Utilities.newBlob(
        Utilities.base64Decode(base64Data),
        incData.fileObj.mimeType,
        `${incidentId}_${incData.fileObj.fileName}`,
      );
      const file = folder.createFile(blob);
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
      attachmentUrl = file.getUrl();
    }

    incidentSheet.appendRow([
      incidentId,
      incData.type,
      incData.location,
      incData.description || "",
      incData.reporterEmail,
      "In Progress",
      attachmentUrl,
      createdAt,
    ]);

    sendLineNotification(
      incData,
      incidentId,
      attachmentUrl,
      createdAt,
      reporterDetail,
    );

    return { success: true, incidentId: incidentId };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getIncidents(userEmail) {
  try {
    const { incidentSheet } = getDb();
    const data = incidentSheet.getDataRange().getValues();
    const results = [];
    const isAdmin = checkIsAdmin(userEmail);
    const searchEmail = (userEmail || "").toString().trim().toLowerCase();

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (!row[0]) continue;

      const reporter = (row[4] || "").toString().trim().toLowerCase();

      if (isAdmin || (searchEmail && reporter === searchEmail)) {
        const createdAtStr = row[7]
          ? row[7] instanceof Date
            ? Utilities.formatDate(
                row[7],
                "Asia/Bangkok",
                "yyyy-MM-dd HH:mm:ss",
              )
            : row[7].toString()
          : "";
        results.push({
          IncidentID: row[0].toString(),
          Type: row[1] ? row[1].toString() : "",
          Location: row[2] ? row[2].toString() : "",
          Description: row[3] ? row[3].toString() : "",
          ReporterEmail: row[4] ? row[4].toString() : "",
          Status: row[5] ? row[5].toString() : "",
          AttachmentUrl: row[6] ? row[6].toString() : "",
          CreatedAt: createdAtStr,
        });
      }
    }
    return { success: true, data: results };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateIncidentStatus(incidentId, newStatus) {
  try {
    const { incidentSheet } = getDb();
    const data = incidentSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === incidentId.toString()) {
        incidentSheet.getRange(i + 1, 6).setValue(newStatus);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบรหัสเหตุการณ์" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ----------------- ADMIN USER MANAGEMENT -----------------
function getAllUsers(adminEmail) {
  if (!checkIsAdmin(adminEmail)) {
    return {
      success: false,
      message: "คุณไม่มีสิทธิ์เข้าถึงข้อมูลผู้ใช้ (เฉพาะ Admin เท่านั้น)",
    };
  }

  try {
    const { userSheet } = getDb();
    const data = userSheet.getDataRange().getValues();
    const users = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      users.push({
        Email: row[0].toString(),
        Name: row[2] ? row[2].toString() : "",
        ID: row[3] ? row[3].toString() : "",
        Group: row[4] ? row[4].toString() : "",
        Faculty: row[5] ? row[5].toString() : "",
        Department: row[6] ? row[6].toString() : "",
        Role: row[7] ? row[7].toString() : "",
        Status: row[8] ? row[8].toString() : "",
      });
    }
    return { success: true, data: users };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateUserStatus(adminEmail, email, newStatus) {
  if (!checkIsAdmin(adminEmail)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const { userSheet } = getDb();
    const data = userSheet.getDataRange().getValues();
    const emailLower = email.toString().trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === emailLower) {
        userSheet.getRange(i + 1, 9).setValue(newStatus);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบอีเมลผู้ใช้" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ----------------- STATS & TICKER -----------------
function getIncidentStats(userEmail) {
  if (!checkIsAdmin(userEmail)) {
    return {
      success: false,
      message: "คุณไม่มีสิทธิ์เข้าถึงข้อมูลสถิติ (เฉพาะ Admin เท่านั้น)",
    };
  }

  try {
    const { incidentSheet } = getDb();
    const data = incidentSheet.getDataRange().getValues();

    let total = 0,
      active = 0,
      resolved = 0;
    const typeCount = {};
    const monthCount = new Array(12).fill(0);
    const yearCount = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      total++;
      const status = row[5] ? row[5].toString() : "";
      if (status === "In Progress" || status === "กำลังดำเนินการ") active++;
      else if (status === "Resolved" || status === "เสร็จสิ้น") resolved++;

      const type = row[1] ? row[1].toString() : "อื่นๆ";
      typeCount[type] = (typeCount[type] || 0) + 1;

      if (row[7]) {
        const date = new Date(row[7]);
        if (!isNaN(date.getTime())) {
          monthCount[date.getMonth()]++;
          const y = date.getFullYear().toString();
          yearCount[y] = (yearCount[y] || 0) + 1;
        }
      }
    }

    return {
      success: true,
      total: total,
      active: active,
      resolved: resolved,
      typeCount: typeCount,
      monthCount: monthCount,
      yearCount: yearCount,
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getUserFacultyStats(userEmail) {
  if (!checkIsAdmin(userEmail)) {
    return {
      success: false,
      message: "คุณไม่มีสิทธิ์เข้าถึงข้อมูลสถิติ (เฉพาะ Admin เท่านั้น)",
    };
  }

  try {
    const { userSheet } = getDb();
    const data = userSheet.getDataRange().getValues();
    const facultyCount = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      const faculty = row[5] ? row[5].toString().trim() : "ไม่ระบุคณะ";
      facultyCount[faculty] = (facultyCount[faculty] || 0) + 1;
    }

    return { success: true, facultyCount: facultyCount };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getLatestIncident() {
  try {
    const { incidentSheet } = getDb();
    const data = incidentSheet.getDataRange().getValues();

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const status = row[5] ? row[5].toString() : "";
      if (row[0] && (status === "In Progress" || status === "กำลังดำเนินการ")) {
        const createdAtStr = row[7]
          ? row[7] instanceof Date
            ? Utilities.formatDate(
                row[7],
                "Asia/Bangkok",
                "yyyy-MM-dd HH:mm:ss",
              )
            : row[7].toString()
          : "";
        return {
          success: true,
          data: {
            IncidentID: row[0].toString(),
            Type: row[1] ? row[1].toString() : "",
            Location: row[2] ? row[2].toString() : "",
            Description: row[3] ? row[3].toString() : "",
            CreatedAt: createdAtStr,
          },
        };
      }
    }
    return { success: true, data: null };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ----------------- LINE NOTIFICATION -----------------
function sendLineNotification(
  incData,
  incidentId,
  attachmentUrl,
  createdAt,
  reporterDetail,
) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("LINE_CHANNEL_ACCESS_TOKEN");
  const targetId = props.getProperty("LINE_TARGET_ID");

  if (!token || !targetId) {
    Logger.log(
      "ไม่พบ LINE_CHANNEL_ACCESS_TOKEN หรือ LINE_TARGET_ID ใน Script Properties",
    );
    return;
  }

  const messageText =
    `🚨 แจ้งเหตุฉุกเฉินใหม่!\n` +
    `🆔 รหัสเหตุการณ์: ${incidentId}\n` +
    `📅 เวลาที่แจ้ง: ${createdAt}\n` +
    `📌 ประเภท: ${incData.type}\n` +
    `📍 สถานที่: ${incData.location}\n` +
    `📝 รายละเอียด: ${incData.description || "-"}\n` +
    `👤 ผู้แจ้ง: ${reporterDetail}` +
    (attachmentUrl ? `\nหลักฐาน/ไฟล์แนบ: ${attachmentUrl}` : "");

  const payload = {
    to: targetId,
    messages: [
      {
        type: "text",
        text: messageText,
      },
    ],
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + token,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", options);
  } catch (err) {
    Logger.log("LINE Push Message Error: " + err.toString());
  }
}
