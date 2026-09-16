/**
 * كود Google Apps Script لإدارة استمارة النشاطات والربط بجوجل درايف
 * هذا الكود يربط تطبيق الويب بملف Google Sheets وجوجل درايف (Google Drive) الخاص بك.
 * 
 * طريقة الإعداد والنشر:
 * 1. افتح جدول بيانات Google (Google Sheets) الخاص بك.
 * 2. من القائمة العلوية، اختر: الامتدادات (Extensions) -> Apps Script.
 * 3. امسح أي كود موجود في المحرر، ثم الصق هذا الكود بالكامل.
 * 4. اضغط على زر حفظ (أيقونة القرص).
 * 5. اضغط على زر "نشر" (Deploy) -> "نشر جديد" (New deployment).
 * 6. اختر نوع النشر: "تطبيق ويب" (Web app).
 * 7. اضبط الإعدادات التالية:
 *    - الوصف: قاعدة بيانات النشاطات والربط بجوجل درايف
 *    - تنفيذ التطبيق باسم (Execute as): "أنا" (Me - بريدك الإلكتروني).
 *    - من لديه صلاحية الوصول (Who has access): "أي شخص" (Anyone).
 * 8. اضغط على "نشر" (Deploy)، ثم قم بالموافقة على الصلاحيات المطلوبة (Authorize access) والسماح للتطبيق بالوصول إلى Google Drive و Google Sheets.
 * 9. انسخ "رابط تطبيق الويب" (Web App URL) الناتج، والصقه في إعدادات الأدمن بالتطبيق.
 */

// البريد الإلكتروني الذي ستصل إليه نسخ الاستمارات. 
var ADMIN_EMAIL = "safaa692003@gmail.com"; 

function doGet(e) {
  var sheet = getOrCreateSheet();
  var data = sheet.getDataRange().getValues();
  var records = [];
  
  if (data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // تخطي الصفوف الفارغة
      try {
        records.push({
          id: row[0].toString(),
          timestamp: Number(row[1]),
          date: row[2].toString(),
          month: row[3].toString(),
          year: row[4].toString(),
          employee: row[5].toString(),
          workplace: row[6].toString(),
          activities: JSON.parse(row[7]),
          savedBy: row[8] ? row[8].toString() : "",
          expiryDate: row[9] ? Number(row[9]) : 0,
          lastUpdated: row[10] ? row[10].toString() : "",
          updatedBy: row[11] ? row[11].toString() : ""
        });
      } catch (err) {
        console.error("خطأ في قراءة الصف: " + (i + 1), err);
      }
    }
  }
  
  // ترتيب تنازلي حسب وقت الحفظ
  records.sort(function(a, b) {
    return (b.timestamp || 0) - (a.timestamp || 0);
  });
  
  // جلب المستخدمين أيضاً
  var users = getUsersFromSheet();
  var responseData = {
    records: records,
    users: users
  };
  
  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "No data sent" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var postData = JSON.parse(e.postData.contents);
    
    // حالة حفظ/تعديل مستخدم
    if (postData.action === 'saveUser') {
      var sheet = getOrCreateUsersSheet();
      var data = sheet.getDataRange().getValues();
      var username = postData.username;
      var password = postData.password;
      var role = postData.role;
      var createdBy = postData.createdBy || "system";
      var foundRowIndex = -1;
      
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString() === username.toString()) {
          foundRowIndex = i + 1;
          break;
        }
      }
      
      var rowValues = [
        username,
        password,
        role,
        new Date().toISOString(),
        createdBy
      ];
      
      if (foundRowIndex !== -1) {
        // تحديث كلمة المرور أو الصلاحية
        sheet.getRange(foundRowIndex, 2).setValue(password);
        sheet.getRange(foundRowIndex, 3).setValue(role);
      } else {
        // إضافة مستخدم جديد
        sheet.appendRow(rowValues);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "تم حفظ حساب المستخدم بنجاح" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // حالة حذف مستخدم
    if (postData.action === 'deleteUser') {
      var sheet = getOrCreateUsersSheet();
      var data = sheet.getDataRange().getValues();
      var usernameToDelete = postData.username;
      
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString() === usernameToDelete.toString()) {
          sheet.deleteRow(i + 1);
          return ContentService.createTextOutput(JSON.stringify({ success: true, message: "تم حذف المستخدم بنجاح" }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "المستخدم غير موجود" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // حالة الحذف
    if (postData.action === 'delete') {
      var sheet = getOrCreateSheet();
      var data = sheet.getDataRange().getValues();
      var idToDelete = postData.id;
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString() === idToDelete.toString()) {
          sheet.deleteRow(i + 1); // الصفوف في شيت تبدأ من 1 (العناوين هي الصف 1)
          return ContentService.createTextOutput(JSON.stringify({ success: true, message: "تم الحذف بنجاح" }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "السجل غير موجود" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // حالة الحفظ أو التعديل
    var record = postData;
    var recordId = record.id;
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();
    var foundRowIndex = -1;
    
    if (recordId) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString() === recordId.toString()) {
          foundRowIndex = i + 1;
          break;
        }
      }
    }
    
    var rowValues = [
      record.id,
      record.timestamp,
      record.date,
      record.month,
      record.year,
      record.employee,
      record.workplace,
      JSON.stringify(record.activities || []),
      record.savedBy || "",
      record.expiryDate || 0,
      record.lastUpdated || "",
      record.updatedBy || ""
    ];
    
    if (foundRowIndex !== -1) {
      // تحديث الصف الحالي في جوجل شيت لمنع التكرار ودعم الحفظ التراكمي
      sheet.getRange(foundRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      // إضافة صف جديد
      sheet.appendRow(rowValues);
    }
    
    // 1. إرسال إشعار بريد إلكتروني فوري للمسؤول (مستقل عن درايف)
    try {
      sendEmailNotification(record);
    } catch (emailErr) {
      console.error("خطأ في إرسال البريد الإلكتروني للمسؤول: ", emailErr);
    }

    // 2. حفظ وتحديث نسخة في مجلد Google Drive (مستقل عن البريد)
    try {
      saveToGoogleDrive(record);
    } catch (driveErr) {
      console.error("فشل حفظ الملف في جوجل درايف: ", driveErr);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "تم الحفظ والتحديث بنجاح" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    var headers = ["username", "password", "role", "createdAt", "createdBy"];
    sheet.appendRow(headers);
    // تنسيق العناوين بشكل احترافي
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#c0392b")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    
    // إضافة الحسابات الافتراضية
    var defaultUsers = [
      ["admin", "SA@@2026##ra", "admin", new Date().toISOString(), "system"],
      ["user", "user@159", "user", new Date().toISOString(), "system"],
      ["hossin.f", "hos@789", "user", new Date().toISOString(), "system"],
      ["taher.t", "taher@123", "user", new Date().toISOString(), "system"],
      ["fatima.a", "fatima#753", "user", new Date().toISOString(), "system"],
      ["hamza.j", "hamza@159", "user", new Date().toISOString(), "system"],
      ["staar.a", "sataar753", "user", new Date().toISOString(), "system"],
      ["drgham", "drgham@@753", "user", new Date().toISOString(), "system"]
    ];
    defaultUsers.forEach(function(u) {
      sheet.appendRow(u);
    });
  }
  return sheet;
}

function getUsersFromSheet() {
  var sheet = getOrCreateUsersSheet();
  var data = sheet.getDataRange().getValues();
  var users = {};
  if (data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      users[row[0].toString()] = {
        password: row[1].toString(),
        role: row[2].toString(),
        createdAt: row[3].toString(),
        createdBy: row[4].toString()
      };
    }
  }
  return users;
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Records");
  if (!sheet) {
    sheet = ss.insertSheet("Records");
    var headers = [
      "id", "timestamp", "date", "month", "year", "employee", "workplace", 
      "activities", "savedBy", "expiryDate", "lastUpdated", "updatedBy"
    ];
    sheet.appendRow(headers);
    // تنسيق العناوين بشكل احترافي
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1e3c72")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
  }
  return sheet;
}

/**
 * دالة لحفظ وتحديث نسخة من الاستمارة على شكل جدول بيانات Google Sheets داخل مجلد بجوجل درايف
 */
function saveToGoogleDrive(record) {
  var folderName = "استمارات النشاطات الشهرية";
  var folders = DriveApp.getFoldersByName(folderName);
  var folder;
  
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
  
  // اسم الملف: اسم الموظف - الشهر السنة - مكان العمل (بدون ملحق .txt)
  var fileName = record.employee + " - " + record.month + " " + record.year + " - " + record.workplace;
  
  // البحث عن الملف الحالي إذا كان موجوداً لتعديله (حفظ تراكمي) أو إنشاء ملف جديد
  var files = folder.getFilesByName(fileName);
  var file;
  var ss;
  if (files.hasNext()) {
    file = files.next();
    ss = SpreadsheetApp.openById(file.getId());
  } else {
    ss = SpreadsheetApp.create(fileName);
    file = DriveApp.getFileById(ss.getId());
    file.moveTo(folder);
  }
  
  var sheet = ss.getSheets()[0];
  sheet.clear(); // تفريغ المحتوى والتنسيقات القديمة لإعادة الكتابة بشكل منسق
  sheet.setRightToLeft(true);
  
  // 1. الترويسة الرئيسية
  sheet.getRange("A1:I1").merge().setValue("استمارة النشاطات الشهرية - قسم الاعداد والتدريب")
       .setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center").setBackground("#1e3c72").setFontColor("#ffffff");
  sheet.getRange("A2:I2").merge().setValue("المديرية العامة لتربية المثنى - وزارة التربية")
       .setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center").setBackground("#2a5298").setFontColor("#ffffff");
  
  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 25);
  
  // 2. معلومات الاستمارة
  var metaData = [
    ["الموظف المسؤول:", record.employee, "", "مكان العمل:", record.workplace, "", "", "", ""],
    ["الشهر / السنة:", record.month + " " + record.year, "", "تاريخ الإرسال:", record.date, "", "", "", ""]
  ];
  sheet.getRange("A4:I5").setValues(metaData);
  sheet.getRange("A4:A5").setFontWeight("bold").setBackground("#f0f4f8");
  sheet.getRange("D4:D5").setFontWeight("bold").setBackground("#f0f4f8");
  
  // 3. جدول الأنشطة - العناوين
  var headers = ["ت", "اسم النشاط", "التفاصيل", "المشاركين", "جهة التنفيذ", "الجهة المستهدفة", "المرفقات", "التاريخ", "الملاحظات"];
  sheet.getRange("A7:I7").setValues([headers])
       .setFontWeight("bold").setBackground("#2c3e50").setFontColor("#ffffff").setHorizontalAlignment("center");
  sheet.setRowHeight(7, 25);
  
  // 4. كتابة صفوف الأنشطة
  var startRow = 8;
  var activities = record.activities || [];
  var totalParticipants = 0;
  
  if (activities.length > 0) {
    var rowsData = [];
    activities.forEach(function(act, idx) {
      var p = parseInt(act.participants) || 0;
      totalParticipants += p;
      rowsData.push([
        idx + 1,
        act.name || "",
        act.details || "",
        p,
        act.executor || "",
        act.target || "",
        act.attachments || "",
        act.date || "",
        act.notes || ""
      ]);
    });
    sheet.getRange(startRow, 1, rowsData.length, 9).setValues(rowsData)
         .setHorizontalAlignment("center").setVerticalAlignment("middle");
    
    // إضافة حدود للجدول
    sheet.getRange(7, 1, rowsData.length + 1, 9).setBorder(true, true, true, true, true, true);
    
    // سطر المجموع
    var sumRow = startRow + rowsData.length;
    sheet.getRange(sumRow, 1, 1, 3).merge().setValue("إجمالي المشاركين في جميع الأنشطة")
         .setFontWeight("bold").setBackground("#eaeded").setHorizontalAlignment("center");
    sheet.getRange(sumRow, 4).setValue(totalParticipants)
         .setFontWeight("bold").setBackground("#27ae60").setFontColor("#ffffff").setHorizontalAlignment("center");
    sheet.getRange(sumRow, 5, 1, 5).merge().setValue("")
         .setBackground("#eaeded");
    sheet.getRange(sumRow, 1, 1, 9).setBorder(true, true, true, true, true, true);
    sheet.setRowHeight(sumRow, 25);
    
    startRow = sumRow + 2; // لنبدأ منطقة التواقيع بعد سطر المجموع
  } else {
    sheet.getRange("A8:I8").merge().setValue("لا توجد أنشطة مدخلة في هذه الاستمارة.")
         .setHorizontalAlignment("center").setFontStyle("italic");
    sheet.getRange(7, 1, 2, 9).setBorder(true, true, true, true, true, true);
    startRow = 10;
  }
  
  // 5. قسم التوقيعات
  sheet.getRange(startRow, 2).setValue("التوقيع :").setFontWeight("bold");
  sheet.getRange(startRow, 7).setValue("التوقيع :").setFontWeight("bold");
  
  sheet.getRange(startRow + 1, 2).setValue("مسؤول الشعبة").setFontWeight("bold");
  sheet.getRange(startRow + 1, 7).setValue("مدير القسم").setFontWeight("bold");
  
  sheet.getRange(startRow + 2, 7).setValue("م.م صفاء هادي عبد العالي").setFontWeight("bold");
  
  sheet.getRange(startRow, 2, 3, 1).setHorizontalAlignment("center");
  sheet.getRange(startRow, 7, 3, 1).setHorizontalAlignment("center");
  
}

/**
 * دالة لتوليد المحتوى النصي للاستمارة
 */
function buildRecordTextContent(record) {
  var textContent = "==================================================\n";
  textContent += "      استمارة النشاطات الشهرية - قسم الاعداد والتدريب\n";
  textContent += "        المديرية العامة لتربية المثنى - وزارة التربية\n";
  textContent += "==================================================\n\n";
  textContent += "الموظف المسؤول: " + record.employee + "\n";
  textContent += "مكان العمل:     " + record.workplace + "\n";
  textContent += "الشهر / السنة:  " + record.month + " " + record.year + "\n";
  textContent += "تاريخ الإرسال:   " + record.date + "\n";
  if (record.lastUpdated) {
    textContent += "آخر تحديث:      " + record.lastUpdated + " (بواسطة: " + (record.updatedBy || "") + ")\n";
  }
  textContent += "إجمالي الأنشطة: " + (record.activities ? record.activities.length : 0) + "\n\n";
  textContent += "--------------------------------------------------\n";
  textContent += "تفاصيل الأنشطة في الاستمارة:\n";
  textContent += "--------------------------------------------------\n";
  
  if (record.activities && record.activities.length > 0) {
    record.activities.forEach(function(act, idx) {
      textContent += (idx + 1) + ". اسم النشاط:  " + (act.name || "") + "\n";
      textContent += "   التفاصيل:    " + (act.details || "") + "\n";
      textContent += "   المشاركين:   " + (act.participants || 0) + " مشارك\n";
      textContent += "   جهة التنفيذ:  " + (act.executor || "") + "\n";
      textContent += "   الجهة المستهدفة: " + (act.target || "") + "\n";
      textContent += "   المرفقات:    " + (act.attachments || "") + "\n";
      textContent += "   التاريخ:      " + (act.date || "") + "\n";
      textContent += "   الملاحظات:   " + (act.notes || "") + "\n";
      textContent += "--------------------------------------------------\n";
    });
  } else {
    textContent += "لا توجد أنشطة مدخلة في هذه الاستمارة.\n";
  }
  
  textContent += "\nتم التوليد تلقائياً بواسطة نظام إدارة الاستمارات - المديرية العامة لتربية المثنى.\n";
  return textContent;
}

/**
 * دالة لتوليد قالب HTML منسق للبريد الإلكتروني
 */
function buildRecordHtmlContent(record) {
  var activities = record.activities || [];
  var totalParticipants = 0;
  var rowsHtml = "";
  
  activities.forEach(function(act, idx) {
    var p = parseInt(act.participants) || 0;
    totalParticipants += p;
    rowsHtml += "<tr style='border-bottom: 1px solid #ddd; text-align: center; font-size: 13px;'>" +
      "<td style='padding: 8px; border: 1px solid #ccc;'>" + (idx + 1) + "</td>" +
      "<td style='padding: 8px; border: 1px solid #ccc; font-weight: bold;'>" + (act.name || "") + "</td>" +
      "<td style='padding: 8px; border: 1px solid #ccc;'>" + (act.details || "") + "</td>" +
      "<td style='padding: 8px; border: 1px solid #ccc; background: #e8f8f5; font-weight: bold;'>" + p + "</td>" +
      "<td style='padding: 8px; border: 1px solid #ccc;'>" + (act.executor || "") + "</td>" +
      "<td style='padding: 8px; border: 1px solid #ccc;'>" + (act.target || "") + "</td>" +
      "<td style='padding: 8px; border: 1px solid #ccc;'>" + (act.attachments || "") + "</td>" +
      "<td style='padding: 8px; border: 1px solid #ccc;'>" + (act.date || "") + "</td>" +
      "<td style='padding: 8px; border: 1px solid #ccc;'>" + (act.notes || "") + "</td>" +
    "</tr>";
  });
  
  if (activities.length === 0) {
    rowsHtml = "<tr><td colspan='9' style='padding: 20px; text-align: center; color: #777;'>لا توجد أنشطة مدخلة</td></tr>";
  }

  var html = "<div dir='rtl' style='font-family: Arial, Tahoma, sans-serif; color: #333; line-height: 1.6; max-width: 800px; margin: auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; background: #fdfdfd;'>" +
    "<div style='background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: #fff; padding: 20px; text-align: center; border-radius: 6px;'>" +
      "<h2 style='margin: 0; font-size: 20px;'>وزارة التربية - المديرية العامة لتربية المثنى</h2>" +
      "<h3 style='margin: 5px 0 0 0; font-size: 16px; font-weight: normal;'>قسم الاعداد والتدريب - استمارة النشاطات الشهرية</h3>" +
    "</div>" +
    
    "<div style='margin: 20px 0; background: #fff; border: 1px solid #eee; border-radius: 6px; padding: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);'>" +
      "<table style='width: 100%; border-collapse: collapse; font-size: 14px;'>" +
        "<tr>" +
          "<td style='padding: 6px; font-weight: bold; width: 20%; color: #1e3c72;'>الموظف المسؤول:</td>" +
          "<td style='padding: 6px; font-weight: bold; width: 30%;'>" + record.employee + "</td>" +
          "<td style='padding: 6px; font-weight: bold; width: 20%; color: #1e3c72;'>مكان العمل:</td>" +
          "<td style='padding: 6px; width: 30%;'>" + record.workplace + "</td>" +
        "</tr>" +
        "<tr>" +
          "<td style='padding: 6px; font-weight: bold; color: #1e3c72;'>الشهر / السنة:</td>" +
          "<td style='padding: 6px;'>" + record.month + " " + record.year + "</td>" +
          "<td style='padding: 6px; font-weight: bold; color: #1e3c72;'>تاريخ الإرسال:</td>" +
          "<td style='padding: 6px;'>" + record.date + "</td>" +
        "</tr>" +
        "<tr>" +
          "<td style='padding: 6px; font-weight: bold; color: #1e3c72;'>مجموع الأنشطة:</td>" +
          "<td style='padding: 6px;'><span style='background: #3498db; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px;'>" + activities.length + " نشاط</span></td>" +
          "<td style='padding: 6px; font-weight: bold; color: #1e3c72;'>إجمالي المشاركين:</td>" +
          "<td style='padding: 6px;'><span style='background: #27ae60; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px;'>" + totalParticipants + " مشارك</span></td>" +
        "</tr>" +
      "</table>" +
    "</div>" +
    
    "<h4 style='color: #1e3c72; margin-bottom: 10px; border-bottom: 2px solid #1e3c72; padding-bottom: 5px;'>جدول الأنشطة المنجزة:</h4>" +
    "<table style='width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 13px;'>" +
      "<thead>" +
        "<tr style='background: #2c3e50; color: #fff; text-align: center;'>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50; width: 30px;'>#</th>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50;'>النشاط</th>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50;'>التفاصيل</th>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50; width: 60px;'>المشاركين</th>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50;'>جهة التنفيذ</th>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50;'>الجهة المستهدفة</th>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50;'>المرفقات</th>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50; width: 85px;'>التاريخ</th>" +
          "<th style='padding: 8px; border: 1px solid #2c3e50;'>الملاحظات</th>" +
        "</tr>" +
      "</thead>" +
      "<tbody>" + rowsHtml + "</tbody>" +
    "</table>" +
    
    "<div style='margin-top: 30px; text-align: center; color: #888; font-size: 12px; border-top: 1px solid #eee; padding-top: 15px;'>" +
      "تم الإرسال تلقائياً بواسطة نظام إدارة استمارات النشاطات - تصميم وبرمجة: م.م صفاء هادي عبد العالي" +
    "</div>" +
  "</div>";

  return html;
}

/**
 * دالة لإرسال نسخة من الاستمارة عبر البريد الإلكتروني للمسؤول
 */
function sendEmailNotification(record) {
  try {
    var recipientEmail = ADMIN_EMAIL || Session.getEffectiveUser().getEmail();
    
    if (!recipientEmail) {
      console.warn("لم يتم العثور على بريد إلكتروني لإرسال الاستمارة إليه.");
      return;
    }
    
    var subject = "نسخة استمارة: " + record.employee + " - " + record.month + " " + record.year + " - " + record.workplace;
    var textBody = buildRecordTextContent(record);
    var htmlBody = buildRecordHtmlContent(record);
    
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      body: textBody,
      htmlBody: htmlBody
    });
    
    console.log("تم إرسال نسخة من الاستمارة بنجاح إلى: " + recipientEmail);
  } catch (emailErr) {
    console.error("فشل إرسال البريد الإلكتروني: ", emailErr);
  }
}
