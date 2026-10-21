var NOTIFY_EMAIL = PropertiesService.getScriptProperties().getProperty("NOTIFY_EMAIL") || "";
var CANCELLED_PREFIX = "[취소] ";
var SERVICE_VERSION = "2026-10-21-direct-management-v2";
var RESERVATION_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty("RESERVATION_SPREADSHEET_ID") || "";
var RESERVATION_SHEET_NAME = "시트1";
var RESERVATION_PAGE_URL = "https://birkits-kor.github.io/showroom/";
var SHOWROOM_MAP_URLS = {
  "삼성점(봉은사역)": "https://naver.me/FbVBWSQz",
  "압구정점": "https://naver.me/5PVXL3GU",
  "아주대 병원점(수원시 영통구)": "https://naver.me/FY3Odocw",
  "회기점(동대문구)": "https://naver.me/FK0NVfhW",
  "이브자리 양재(하나로마트 정문 우측)": "https://naver.me/FRoQ1AYk",
  "파주운정점(스타필드 빌리지)": "https://naver.me/Ffsewkgu"
};
var SHOWROOMS = [
  "삼성점(봉은사역)",
  "압구정점",
  "아주대 병원점(수원시 영통구)",
  "회기점(동대문구)",
  "이브자리 양재(하나로마트 정문 우측)",
  "파주운정점(스타필드 빌리지)"
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || "{}");
    var action = data.action || "createReservation";
    if (/^managed(Get|Update|Cancel)$/.test(action)) return jsonResponse(managementHandle_(data));
    managementSecret_();

    if (action === "createReservation") return createReservation(data);
    if (action === "lookupReservation") return lookupReservation(data);
    if (action === "updateReservation") return updateReservation(data);
    if (action === "cancelReservation") return cancelReservation(data);

    return jsonResponse({ result: "error", message: "지원하지 않는 요청입니다." });
  } catch (error) {
    return jsonResponse({ result: "error", message: error.message });
  }
}

function doGet(e) {
  try {
    if (e.parameter.action === "getServiceInfo") {
      return jsonResponse({ result: "success", version: SERVICE_VERSION });
    }
    if (e.parameter.action !== "getBookedTimes") {
      return jsonResponse({ result: "error", message: "지원하지 않는 요청입니다." });
    }

    var sheet = getReservationSheet();
    var targetDate = cleanText(e.parameter.date);
    var targetShowroom = cleanText(e.parameter.showroom);
    var excludeRowId = Number(e.parameter.excludeRowId || 0);
    var values = sheet.getDataRange().getValues();
    var bookedTimes = [];

    for (var i = 1; i < values.length; i++) {
      var rowNumber = i + 1;
      var rowShowroom = cleanText(values[i][3]);
      if (rowNumber === excludeRowId || isCancelledReservation(values[i])) continue;

      var parsed = parseDateTimeCell(values[i][4], sheet);
      if (rowShowroom === targetShowroom && parsed.date === targetDate && parsed.time) {
        bookedTimes.push(parsed.time);
      }
    }

    return jsonResponse({ result: "success", bookedTimes: bookedTimes });
  } catch (error) {
    return jsonResponse({ result: "error", message: error.message, bookedTimes: [] });
  }
}

function createReservation(data) {
  var validationMessage = validateReservationData(data);
  if (validationMessage) return jsonResponse({ result: "error", message: validationMessage });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet = getReservationSheet();
    ensureReservationColumns(sheet);
    if (hasSlotConflict(sheet, data.showroom, data.date, data.time, 0)) {
      return jsonResponse({ result: "error", code: "slot_unavailable", message: "방금 다른 예약이 접수되어 선택한 시간이 마감되었습니다. 다른 시간을 선택해 주세요." });
    }

    sheet.appendRow([
      new Date(),
      safeSheetText(data.name),
      safeSheetText(data.phone),
      safeSheetText(data.showroom),
      safeSheetText(data.date + " " + data.time),
      safeSheetText(data.memo || ""),
      safeSheetText(normalizeEmail(data.email)),
      "예약", "", "", new Date(), ""
    ]);
    setNoShowCountFormula(sheet, sheet.getLastRow());
    data._managementRow=sheet.getLastRow();
    data._managementToken=managementIssue_(sheet,data._managementRow,data.email);

    var customerEmailSent = sendReservationNotifications("created", data);

    return jsonResponse({ result: "success", action: "created", customerEmailSent: customerEmailSent });
  } finally {
    lock.releaseLock();
  }
}

function lookupReservation(data) {
  var name = cleanText(data.name);
  var phone = normalizePhone(data.phone);
  if (!name || phone.length < 9) {
    return jsonResponse({ result: "error", message: "예약자 성함과 연락처를 정확히 입력해 주세요." });
  }

  var sheet = getReservationSheet();
  var values = sheet.getDataRange().getValues();
  var nowText = Utilities.formatDate(new Date(), getSheetTimeZone(sheet), "yyyy-MM-dd HH:mm");
  var reservations = [];

  for (var i = 1; i < values.length; i++) {
    var rowName = cleanText(values[i][1]);
    var rowPhone = normalizePhone(values[i][2]);
    var rowShowroom = cleanText(values[i][3]);
    var parsed = parseDateTimeCell(values[i][4], sheet);

    var reservationDateTime = parsed.date && parsed.time ? parsed.date + " " + parsed.time : "";
    if (rowName === name && rowPhone === phone && !isCancelledReservation(values[i]) && getReservationStatus(values[i]) === "예약" && reservationDateTime >= nowText) {
      reservations.push({
        id: String(i + 1),
        showroom: rowShowroom,
        date: parsed.date,
        time: parsed.time,
        memo: cleanText(values[i][5]),
        email: normalizeEmail(values[i][6])
      });
    }
  }

  reservations.sort(function(a, b) {
    return (a.date + " " + a.time).localeCompare(b.date + " " + b.time);
  });

  return jsonResponse({ result: "success", reservations: reservations });
}

function updateReservation(data) {
  var validationMessage = validateReservationData(data);
  if (validationMessage) return jsonResponse({ result: "error", message: validationMessage });

  var rowNumber = Number(data.reservationId);
  var lookupName = cleanText(data.lookupName);
  var lookupPhone = normalizePhone(data.lookupPhone);
  if (!Number.isInteger(rowNumber) || rowNumber < 2 || !lookupName || lookupPhone.length < 9) {
    return jsonResponse({ result: "error", message: "예약 확인 정보가 올바르지 않습니다. 다시 조회해 주세요." });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet = getReservationSheet();
    ensureReservationColumns(sheet);
    var reservation = verifyReservationOwner(sheet, rowNumber, lookupName, lookupPhone);
    if (!reservation) {
      return jsonResponse({ result: "error", message: "예약을 찾을 수 없습니다. 다시 조회해 주세요." });
    }
    if (hasSlotConflict(sheet, data.showroom, data.date, data.time, rowNumber)) {
      return jsonResponse({ result: "error", code: "slot_unavailable", message: "방금 다른 예약이 접수되어 선택한 시간이 마감되었습니다. 다른 시간을 선택해 주세요." });
    }

    var previousReminderKey = reservationReminderKey(reservation);
    sheet.getRange(rowNumber, 4, 1, 4).setValues([[
      safeSheetText(data.showroom),
      safeSheetText(data.date + " " + data.time),
      safeSheetText(data.memo || ""),
      safeSheetText(normalizeEmail(data.email))
    ]]);

    sheet.getRange(rowNumber, 8).setValue("예약");
    sheet.getRange(rowNumber, 11).setValue(new Date());
    if (previousReminderKey !== reservationReminderKey(data)) {
      sheet.getRange(rowNumber, 9, 1, 2).setValues([["", ""]]);
    }
    setNoShowCountFormula(sheet, rowNumber);

    var customerEmailSent = sendReservationNotifications("updated", {
      name: reservation.name, phone: reservation.phone, email: normalizeEmail(data.email), _managementRow:rowNumber,
      showroom: data.showroom, date: data.date, time: data.time, memo: data.memo
    }, reservation);

    return jsonResponse({ result: "success", action: "updated", customerEmailSent: customerEmailSent });
  } finally {
    lock.releaseLock();
  }
}

function cancelReservation(data) {
  var rowNumber = Number(data.reservationId);
  var lookupName = cleanText(data.lookupName);
  var lookupPhone = normalizePhone(data.lookupPhone);
  if (!Number.isInteger(rowNumber) || rowNumber < 2 || !lookupName || lookupPhone.length < 9) {
    return jsonResponse({ result: "error", message: "예약 확인 정보가 올바르지 않습니다. 다시 조회해 주세요." });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet = getReservationSheet();
    var reservation = verifyReservationOwner(sheet, rowNumber, lookupName, lookupPhone);
    if (!reservation) {
      return jsonResponse({ result: "error", message: "예약을 찾을 수 없습니다. 다시 조회해 주세요." });
    }

    var cancelledAt = Utilities.formatDate(new Date(), getSheetTimeZone(sheet), "yyyy-MM-dd HH:mm");
    var cancellationNote = "[" + cancelledAt + " 온라인 취소]";
    var updatedMemo = reservation.memo ? reservation.memo + "\n" + cancellationNote : cancellationNote;

    sheet.getRange(rowNumber, 4).setValue(CANCELLED_PREFIX + reservation.showroom);
    sheet.getRange(rowNumber, 6).setValue(safeSheetText(updatedMemo));
    sheet.getRange(rowNumber, 8).setValue("취소");

    reservation._managementRow=rowNumber;
    var customerEmailSent = sendReservationNotifications("cancelled", reservation);

    return jsonResponse({ result: "success", action: "cancelled", customerEmailSent: customerEmailSent });
  } finally {
    lock.releaseLock();
  }
}

function verifyReservationOwner(sheet, rowNumber, name, phone) {
  if (!Number.isInteger(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) return null;

  var values = sheet.getRange(rowNumber, 1, 1, 12).getValues()[0];
  var rowName = cleanText(values[1]);
  var rowPhone = normalizePhone(values[2]);
  var rowShowroom = cleanText(values[3]);
  if (rowName !== name || rowPhone !== phone || getReservationStatus(values) !== "예약") return null;

  var parsed = parseDateTimeCell(values[4], sheet);
  var nowText = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");
  if (!parsed.date || !parsed.time || parsed.date + " " + parsed.time < nowText) return null;
  return {
    name: rowName,
    phone: cleanText(values[2]),
    showroom: rowShowroom,
    date: parsed.date,
    time: parsed.time,
    memo: cleanText(values[5]),
    email: normalizeEmail(values[6])
  };
}

function hasSlotConflict(sheet, showroom, date, time, excludeRowNumber) {
  var values = sheet.getDataRange().getValues();
  var hour = time.split(":")[0];
  var legacyHalfHour = hour + ":30";

  for (var i = 1; i < values.length; i++) {
    var rowNumber = i + 1;
    var rowShowroom = cleanText(values[i][3]);
    if (rowNumber === excludeRowNumber || isCancelledReservation(values[i])) continue;

    var parsed = parseDateTimeCell(values[i][4], sheet);
    if (rowShowroom === showroom && parsed.date === date && (parsed.time === time || parsed.time === legacyHalfHour)) {
      return true;
    }
  }
  return false;
}

function validateReservationData(data) {
  var name = cleanText(data.name);
  var phone = normalizePhone(data.phone);
  var email = normalizeEmail(data.email);
  var showroom = cleanText(data.showroom);
  var date = cleanText(data.date);
  var time = cleanText(data.time);

  if (data.visitPromiseAccepted !== true) return "방문 약속 안내를 확인하고 필수 항목에 체크해 주세요.";
  if (!name || phone.length < 9) return "성함과 연락처를 정확히 입력해 주세요.";
  if (!isValidEmail(email)) return "예약 안내를 받으실 이메일 주소를 정확히 입력해 주세요.";
  if (SHOWROOMS.indexOf(showroom) === -1) return "방문하실 매장을 다시 선택해 주세요.";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:00$/.test(time)) return "방문 날짜와 1시간 단위 시간을 다시 선택해 주세요.";
  if (!isAllowedReservationTime(showroom, date, time)) return "선택한 일정은 해당 매장의 예약 가능 시간이 아닙니다.";
  return "";
}

function isAllowedReservationTime(showroom, date, time) {
  var dateParts = date.split("-");
  var dateObject = new Date(Number(dateParts[0]), Number(dateParts[1]) - 1, Number(dateParts[2]));
  if (isNaN(dateObject.getTime()) ||
      Utilities.formatDate(dateObject, "Asia/Seoul", "yyyy-MM-dd") !== date) return false;

  var selectedHour = Number(time.split(":")[0]);
  var selectedDateTime = new Date(Number(dateParts[0]), Number(dateParts[1]) - 1, Number(dateParts[2]), selectedHour, 0, 0, 0);
  if (selectedDateTime <= new Date()) return false;

  var day = dateObject.getDay();
  var holiday = isFixedHoliday(dateObject);
  var openHour = 10;
  var closeHour = 20;

  if (showroom === "삼성점(봉은사역)") {
    closeHour = (day >= 1 && day <= 5) ? 21 : 20.5; // 월~금 21:00, 토~일 20:30
  } else if (showroom === "압구정점") {
    if (day === 0) closeHour = 20.5;
    else if (day === 5) closeHour = 20;
    else closeHour = 21;
  } else if (showroom === "아주대 병원점(수원시 영통구)") {
    if (day === 0 || day === 6 || holiday) return false;
    closeHour = 18;
  } else if (showroom === "회기점(동대문구)") {
    closeHour = day === 0 ? 19 : 21;
  } else if (showroom === "이브자리 양재(하나로마트 정문 우측)") {
    openHour = 9;
    closeHour = 20;
  } else if (showroom === "파주운정점(스타필드 빌리지)") {
    closeHour = 22;
  }

  var selectedMinutes = selectedHour * 60;
  return selectedMinutes >= openHour * 60 && selectedMinutes + 60 <= closeHour * 60;
}

function isFixedHoliday(dateObject) {
  var mm = String(dateObject.getMonth() + 1).padStart(2, "0");
  var dd = String(dateObject.getDate()).padStart(2, "0");
  var fixedHolidays = ["01-01", "03-01", "05-05", "06-06", "08-15", "10-03", "10-09", "12-25"];
  return fixedHolidays.indexOf(mm + "-" + dd) !== -1;
}

function parseDateTimeCell(value, sheet) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return {
      date: Utilities.formatDate(value, getSheetTimeZone(sheet), "yyyy-MM-dd"),
      time: Utilities.formatDate(value, getSheetTimeZone(sheet), "HH:mm")
    };
  }

  var text = cleanText(value);
  var match = text.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2})/);
  if (!match) return { date: "", time: "" };
  return { date: match[1], time: String(match[2]).padStart(2, "0") + ":" + match[3] };
}

function isCancelledShowroom(showroom) {
  return cleanText(showroom).indexOf(CANCELLED_PREFIX) === 0;
}

function normalizePhone(value) {
  return cleanText(value).replace(/\D/g, "");
}

function normalizeEmail(value) {
  return cleanText(value).toLowerCase();
}

function isValidEmail(value) {
  var email = normalizeEmail(value);
  return email.length <= 254 && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function cleanText(value) {
  return value === null || value === undefined ? "" : String(value).trim();
}

function safeSheetText(value) {
  var text = cleanText(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function getReservationSheet() {
  var sheet = SpreadsheetApp.openById(RESERVATION_SPREADSHEET_ID).getSheetByName(RESERVATION_SHEET_NAME);
  if (!sheet) throw new Error("쇼룸 예약 시트를 찾을 수 없습니다.");
  return sheet;
}

function ensureReservationColumns(sheet) {
  var headers = ["이메일", "예약 상태", "전날 안내 발송 일시", "전날 안내 대상 일정", "방문 약속 확인 일시", "노쇼 횟수"];
  var existing = sheet.getRange(1, 7, 1, headers.length).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (cleanText(existing[i]) && cleanText(existing[i]) !== headers[i]) {
      throw new Error("예약 시트 " + (i + 7) + "열 제목이 예상과 다릅니다. 매장에 문의해 주세요.");
    }
  }
  sheet.getRange(1, 7, 1, headers.length).setValues([headers]);
}

function getSheetTimeZone(sheet) {
  return sheet.getParent().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || "Asia/Seoul";
}

function sendNotification(subject, body) {
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL, subject: subject, body: body,
      htmlBody: renderCustomerEmail(body), replyTo: NOTIFY_EMAIL, name: "버킷츠"
    });
    return true;
  } catch (error) {
    console.error("이메일 발송 실패: " + error.message);
    return false;
  }
}

// 접수·수정·취소는 같은 본문과 HTML을 운영자와 고객에게 각각 보냅니다.
function sendReservationNotifications(event, reservation, previous) {
  var status = event === "updated" ? "수정" : event === "cancelled" ? "취소" : "접수";
  var subject = "[버킷츠] 쇼룸 방문 예약 " + status + " 확인";
  var body = buildCustomerMessage(event, reservation, previous);
  sendNotification(subject, body);
  return sendCustomerNotification(reservation.email, subject, body);
}

function sendCustomerNotification(toEmail, subject, body) {
  var email = normalizeEmail(toEmail);
  if (!isValidEmail(email)) return false;

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      htmlBody: renderCustomerEmail(body),
      replyTo: NOTIFY_EMAIL,
      name: "버킷츠"
    });
    return true;
  } catch (error) {
    console.error("고객 이메일 발송 실패: " + error.message);
    return false;
  }
}

function resolveReservationManageUrl(pageUrl) {
  // 사용자 입력 주소 대신 운영자 설정 또는 공식 주소만 메일에 사용합니다.
  var configuredUrl = cleanText(PropertiesService.getScriptProperties().getProperty("RESERVATION_PAGE_URL"));
  var base = /^https:\/\/[a-z0-9.-]+(?::\d+)?(?:[/?#]|$)/i.test(configuredUrl)
    ? configuredUrl : RESERVATION_PAGE_URL;
  return base.split("#")[0].split("?")[0] + "#manageReservation";
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function getReservationStatus(row) {
  if (isCancelledShowroom(row[3]) || cleanText(row[7]) === "취소") return "취소";
  return cleanText(row[7]) || "예약";
}

function isCancelledReservation(row) {
  return getReservationStatus(row) === "취소";
}

function reservationReminderKey(reservation) {
  return [cleanText(reservation.showroom), cleanText(reservation.date), cleanText(reservation.time), normalizeEmail(reservation.email)].join(" | ");
}

function setNoShowCountFormula(sheet, rowNumber) {
  // 연락처 하이픈 차이를 무시하고 매장에서 직접 '노쇼'로 기록한 건만 집계합니다.
  var formula = '=IF(C' + rowNumber + '="","",SUMPRODUCT((REGEXREPLACE(TO_TEXT($C$2:$C),"[^0-9]","")=REGEXREPLACE(TO_TEXT(C' +
    rowNumber + '),"[^0-9]",""))*($H$2:$H="노쇼")))';
  sheet.getRange(rowNumber, 12).setFormula(formula);
}

function reservationMailLinks(reservation) {
  return managementMailLinks_(reservation);
}

function buildCustomerMessage(event, reservation, previous) {
  var links = reservationMailLinks(reservation);
  var title = event === "reminder" ? "내일은 버킷츠 쇼룸 방문 예약일입니다." :
    event === "updated" ? "버킷츠 쇼룸 방문 예약이 변경되었습니다." :
    event === "cancelled" ? "버킷츠 쇼룸 방문 예약이 취소되었습니다." : "버킷츠 쇼룸 방문 예약이 접수되었습니다.";
  var body = title + "\n\n■ 신청 정보\n" +
    "👤 고객명: " + cleanText(reservation.name) + "\n" +
    "📞 연락처: " + cleanText(reservation.phone) + "\n" +
    "✉️ 이메일: " + (normalizeEmail(reservation.email) || "미등록") + "\n";
  if (previous) body += "기존 일정: " + previous.showroom + " / " + previous.date + " " + previous.time + "\n";
  body += "🏢 방문지점: " + reservation.showroom + "\n" +
    "⏰ " + (event === "cancelled" ? "취소 일정" : "방문일정") + ": " + reservation.date + " " + reservation.time + "\n";
  if (event !== "cancelled") body += "상담 및 제품 체험: 약 1시간\n";
  body += "📝 남긴말씀: " + (reservation.memo ? reservation.memo : "없음") + "\n\n";
  var mapUrl = SHOWROOM_MAP_URLS[reservation.showroom];
  if (mapUrl) body += "📍 매장 전화번호 확인·길찾기(네이버 지도):\n" + mapUrl + "\n\n";
  body += "■ 신청 내역 확인·수정·취소\n" +
    "예약 내용 확인:\n" + links.view + "\n\n" +
    "방문 일정 변경:\n" + links.change + "\n\n" +
    "예약 취소:\n" + links.cancel + "\n\n" +
    "본인 신청 전용 보안 링크입니다. 고객 정보 재입력 없이 이용할 수 있으며 다른 사람에게 공유하지 마세요.\n";
  if (event === "cancelled") {
    body += "\n다음에 편하신 일정으로 다시 찾아 주세요.\n";
  } else {
    body += "\n■ 이용 안내\n담당자가 고객님의 예약 시간을 비워두고 체험을 준비합니다.\n" +
      "방문이 어려우시면 다른 고객님이 이용하실 수 있도록 미리 변경하거나 취소해 주세요.\n" +
      "당일 취소 또는 지각 시에는 예약한 매장으로 미리 연락해 주세요.\n";
  }
  return body;
}

function escapeEmailHtml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function renderCustomerEmail(body) {
  return renderBirkitsEmail_(body);
}

// 이 함수는 세 신청 서비스에서 동일하게 유지합니다.

function setupShowroomNoShowPrevention() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getReservationSheet();
    ensureReservationColumns(sheet);
    PropertiesService.getScriptProperties().setProperty("RESERVATION_PAGE_URL", RESERVATION_PAGE_URL);
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (!cleanText(values[i][1])) continue;
      if (!cleanText(values[i][7])) sheet.getRange(i + 1, 8).setValue(getReservationStatus(values[i]));
      if (!sheet.getRange(i + 1, 12).getFormula()) setNoShowCountFormula(sheet, i + 1);
    }
    var validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(["예약", "방문 완료", "취소", "노쇼"], true)
      .setAllowInvalid(false).build();
    sheet.getRange(2, 8, sheet.getMaxRows() - 1, 1).setDataValidation(validation);
    sheet.getRange(1, 8, 1, 5).setFontWeight("bold").setWrap(true).setBackground("#ede8df");
    sheet.setColumnWidth(8, 120);
    sheet.setColumnWidth(9, 180);
    sheet.setColumnWidth(10, 260);
    sheet.setColumnWidth(11, 180);
    sheet.setColumnWidth(12, 90);
    sheet.getRange(2, 9, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm");
    sheet.getRange(2, 11, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm");
    sheet.getRange(1, 8).setNote("방문 후 매장에서 '방문 완료' 또는 '노쇼'로 직접 선택합니다. 시간 경과만으로 노쇼 처리하지 않습니다. '취소'를 선택하면 해당 시간이 다시 예약 가능해집니다.");
    sheet.getRange(1, 9).setNote("전날 안내 메일 발송 성공 시각입니다. 이메일이 없는 기존 예약에는 발송하지 않습니다.");
    sheet.getRange(1, 10).setNote("중복 발송 방지를 위한 매장·일정·이메일 기록입니다. 수동으로 수정하지 마세요.");
    sheet.getRange(1, 11).setNote("온라인 신청·변경 시 필수 방문 약속 확인을 완료한 시각입니다. 기존 예약에는 임의로 채우지 않습니다.");
    sheet.getRange(1, 12).setNote("같은 연락처의 예약 중 매장에서 직접 '노쇼'로 표시한 건수입니다.");
    SpreadsheetApp.flush();
    console.log(JSON.stringify({
      version: SERVICE_VERSION, setup: "complete", reminderTime: "매일 오전 9시경 (Asia/Seoul)",
      reminderInstallation: "Apps Script 트리거 화면에서 매일 오전 9시 실행을 등록합니다.", reminderEmailsSentDuringSetup: 0
    }));
  } finally {
    lock.releaseLock();
  }
}

// 같은 매장·일정·이메일에 성공한 전날 안내는 다시 보내지 않습니다.
function sendTomorrowReservationReminders() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  var summary = { sent: 0, skipped: 0, failed: 0 };
  try {
    var sheet = getReservationSheet();
    ensureReservationColumns(sheet);
    var targetDate = Utilities.formatDate(new Date(Date.now() + 24 * 60 * 60 * 1000), "Asia/Seoul", "yyyy-MM-dd");
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var parsed = parseDateTimeCell(row[4], sheet);
      if (getReservationStatus(row) !== "예약" || parsed.date !== targetDate || !isValidEmail(row[6])) {
        summary.skipped++;
        continue;
      }
      var reservation = {
        name: cleanText(row[1]), phone: cleanText(row[2]), showroom: cleanText(row[3]), date: parsed.date, _managementRow:i+1,
        time: parsed.time, email: normalizeEmail(row[6]), memo: cleanText(row[5])
      };
      var key = reservationReminderKey(reservation);
      if (cleanText(row[9]) === key && row[8]) {
        summary.skipped++;
        continue;
      }
      if (MailApp.getRemainingDailyQuota() < 1) {
        summary.failed++;
        break;
      }
      var sent = sendCustomerNotification(reservation.email,
        "[내일 방문 예정] 버킷츠 " + reservation.showroom + " 쇼룸 예약을 확인해 주세요",
        buildCustomerMessage("reminder", reservation));
      if (sent) {
        sheet.getRange(i + 1, 9, 1, 2).setValues([[new Date(), safeSheetText(key)]]);
        SpreadsheetApp.flush();
        summary.sent++;
      } else summary.failed++;
    }
    PropertiesService.getScriptProperties().setProperty("LAST_REMINDER_RUN",
      JSON.stringify({ at: new Date().toISOString(), targetDate: targetDate, summary: summary }));
    console.log(JSON.stringify(summary));
    return summary;
  } finally {
    lock.releaseLock();
  }
}

// 읽기 전용 점검: 고객 메일 발송이나 예약 변경은 하지 않습니다.
function checkShowroomNoShowPrevention() {
  var sheet = getReservationSheet();
  var result = {
    version: SERVICE_VERSION,
    spreadsheetId: sheet.getParent().getId(),
    sheet: sheet.getName(),
    headers: sheet.getRange(1, 7, 1, 6).getValues()[0],
    lastReminderRun: PropertiesService.getScriptProperties().getProperty("LAST_REMINDER_RUN") || "아직 실행되지 않음"
  };
  console.log(JSON.stringify(result));
  return result;
}

var MANAGEMENT_SERVICE = "showroom";
// 개별 신청 관리. 토큰에는 이름·전화번호·계좌번호를 넣지 않습니다.
function managementConfig_() {
  if (MANAGEMENT_SERVICE === 'showroom') return {base:RESERVATION_PAGE_URL,id:27,email:7,status:8,width:27,headers:{27:'신청 관리 ID'}};
  if (MANAGEMENT_SERVICE === 'moving') return {base:CUSTOMER_PAGE_URL,id:13,email:23,status:28,cancelled:29,paid:30,bank:31,account:32,holder:33,refund:34,width:34,
    headers:{28:'신청 상태',29:'취소 접수일시',30:'입금 여부(고객 신고)',31:'환불 은행',32:'환불 계좌번호',33:'환불 예금주',34:'환불 처리 상태'}};
  return {base:APPLICATION_PAGE_URL,id:27,email:12,status:28,cancelled:29,paid:30,bank:31,account:32,holder:33,refund:34,width:34,
    headers:{27:'신청 관리 ID',28:'신청 상태',29:'취소 접수일시',30:'입금 여부(고객 신고)',31:'환불 은행',32:'환불 계좌번호',33:'환불 예금주',34:'환불 처리 상태'}};
}

function managementSheet_() {
  if (MANAGEMENT_SERVICE === 'showroom') return getReservationSheet();
  if (MANAGEMENT_SERVICE === 'moving') return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  return getUpgradeSheet_();
}

function managementText_(value) { return value == null ? '' : String(value).trim(); }
function managementEmail_(row) { return managementText_(row[managementConfig_().email-1]).toLowerCase(); }
function managementSafeCell_(value) { var text=managementText_(value); return /^[=+\-@]/.test(text)?"'"+text:text; }

function managementEnsureLayout_(sheet) {
  var c=managementConfig_(), max=sheet.getMaxColumns();
  if (max<c.width) sheet.insertColumnsAfter(max,c.width-max);
  Object.keys(c.headers).forEach(function(key) {
    var col=Number(key), current=managementText_(sheet.getRange(1,col).getValue());
    if (current && current!==c.headers[key]) throw new Error('관리용 '+col+'열에 기존 항목이 있습니다. 관리자에게 문의해 주세요.');
    if (!current) {
      var last=sheet.getLastRow();
      if (last>1 && sheet.getRange(2,col,last-1,1).getValues().some(function(r){return managementText_(r[0])!=='';})) throw new Error('관리용 '+col+'열에 기존 데이터가 있습니다.');
      sheet.getRange(1,col).setValue(c.headers[key]);
      sheet.getRange(1,col).setFontWeight('bold').setWrap(true).setBackground('#ede8df');
      sheet.setColumnWidth(col,col===c.account?190:150);
      if (col===c.account) sheet.getRange(2,col,sheet.getMaxRows()-1,1).setNumberFormat('@');
    }
  });
}

// 배포 시 한 번 실행합니다. 고객 기록 변경·메일 발송 없이 열과 비공개 서명 키만 준비합니다.
function initializeDirectManagement() {
  var lock=LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sheet=managementSheet_();
    if (!sheet) throw new Error('대상 신청 시트를 찾을 수 없습니다.');
    managementEnsureLayout_(sheet);
    var props=PropertiesService.getScriptProperties();
    if (!props.getProperty('MANAGEMENT_SIGNING_SECRET')) props.setProperty('MANAGEMENT_SIGNING_SECRET',Utilities.getUuid()+Utilities.getUuid()+Utilities.getUuid());
    SpreadsheetApp.flush();
    var result={service:MANAGEMENT_SERVICE,ready:true,sheet:sheet.getName(),spreadsheetId:sheet.getParent().getId(),emailsSent:0};
    console.log(JSON.stringify(result)); return result;
  } finally {lock.releaseLock();}
}

function managementSecret_() {
  var key=PropertiesService.getScriptProperties().getProperty('MANAGEMENT_SIGNING_SECRET');
  if (!key) throw new Error('개별 관리 링크 설정이 준비되지 않았습니다. 관리자에게 문의해 주세요.');
  return key;
}
function managementEnsureId_(sheet,rowNumber) {
  managementEnsureLayout_(sheet);
  var c=managementConfig_(), id=managementText_(sheet.getRange(rowNumber,c.id).getValue());
  if (!id) {id=Utilities.getUuid();sheet.getRange(rowNumber,c.id).setValue(id);}
  return id;
}
function managementBase64_(value) {return Utilities.base64EncodeWebSafe(value,Utilities.Charset.UTF_8).replace(/=+$/,'');}
function managementSignature_(payload,email) {return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(MANAGEMENT_SERVICE+'|'+payload+'|'+email,managementSecret_(),Utilities.Charset.UTF_8)).replace(/=+$/,'');}
function managementEqual_(a,b) {var x=a.length^b.length;for(var i=0;i<Math.max(a.length,b.length);i++)x|=(a.charCodeAt(i)||0)^(b.charCodeAt(i)||0);return x===0;}
function managementIssue_(sheet,rowNumber,expectedEmail) {
  var c=managementConfig_(), email=managementText_(sheet.getRange(rowNumber,c.email).getValue()).toLowerCase();
  if (expectedEmail!==undefined && email!==managementText_(expectedEmail).toLowerCase()) throw new Error('신청 이메일이 변경되었습니다. 최신 안내 메일을 확인해 주세요.');
  var payload=managementBase64_(JSON.stringify({v:1,id:managementEnsureId_(sheet,rowNumber),exp:Date.now()+180*24*60*60*1000}));
  return payload+'.'+managementSignature_(payload,email);
}
function managementAuthorize_(sheet,token) {
  token=managementText_(token);
  if (token.length>1600 || !/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(token)) throw new Error('관리 링크가 올바르지 않습니다.');
  var parts=token.split('.'), payload;
  try {payload=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString('UTF-8'));} catch(e){throw new Error('관리 링크가 올바르지 않습니다.');}
  if (payload.v!==1 || typeof payload.id!=='string' || !/^[A-Za-z0-9_-]{12,100}$/.test(payload.id) || !Number.isFinite(payload.exp) || payload.exp<Date.now()) throw new Error('관리 링크가 만료되었거나 올바르지 않습니다. 최신 메일을 확인해 주세요.');
  var c=managementConfig_(), last=sheet.getLastRow(), rows=last>1?sheet.getRange(2,1,last-1,c.width).getValues():[], found=0, raw;
  for(var i=0;i<rows.length;i++) if(managementText_(rows[i][c.id-1])===payload.id){if(found)throw new Error('신청 관리 정보가 중복되어 관리자 확인이 필요합니다.');found=i+2;raw=rows[i];}
  if (!found || !managementEqual_(parts[1],managementSignature_(parts[0],managementEmail_(raw)))) throw new Error('관리 링크가 올바르지 않거나 이메일 변경으로 만료되었습니다.');
  return {rowNumber:found,row:raw};
}
function managementMailLinks_(data) {
  var token=data._managementToken;
  if (!token) {
    if (!data._managementRow) throw new Error('개별 신청 관리 정보가 없습니다.');
    token=managementIssue_(managementSheet_(),data._managementRow,data.email===undefined?data.customerEmail:data.email);
  }
  var base=managementConfig_().base.replace(/\/$/,'')+'/manage.html';
  return {view:base+'?action=view#token='+token,change:base+'?action=edit#token='+token,cancel:base+'?action=cancel#token='+token};
}
function managementRefundText_(refund) {
  if (!refund) return '';
  if (!refund.paid) return '\n■ 취소·환불 안내\n입금 전으로 접수되었습니다. 입금 여부는 담당자가 확인합니다.\n';
  var digits=managementText_(refund.account).replace(/\D/g,'');
  return '\n■ 취소·환불 안내\n입금 완료(고객 신고) / 환불 접수 대기\n환불 은행: '+refund.bank+'\n환불 계좌: '+('*'.repeat(Math.max(0,digits.length-4))+digits.slice(-4))+'\n예금주: '+refund.holder+'\n입금 내역 확인 후 담당자가 환불 처리합니다. 자동 송금은 진행되지 않습니다.\n';
}
function managementStatus_(raw) {return MANAGEMENT_SERVICE==='showroom'?getReservationStatus(raw):(managementText_(raw[managementConfig_().status-1])||'신청');}
function managementPublic_(sheet,n,raw) {
  var c=managementConfig_(), status=managementStatus_(raw), data;
  if (MANAGEMENT_SERVICE==='showroom') {
    var parsed=parseDateTimeCell(raw[4],sheet);
    data={name:raw[1],phone:raw[2],email:raw[6]||'',showroom:raw[3],date:parsed.date,time:parsed.time,memo:raw[5]||''};
  } else if(MANAGEMENT_SERVICE==='moving') {
    var display=sheet.getRange(n,1,1,c.width).getDisplayValues()[0];
    data=rowToApplication_(display);
    data.addressBefore=display[14]?(display[14]+' '+display[15]).trim():String(display[4]||'').split(' >> ')[0]||'';
    data.addressAfter=display[16]?(display[16]+' '+display[17]).trim():String(display[4]||'').split(' >> ')[1]||'';
    data.requestDate=display[2]||'';
    var services=['full','assembly','disassembly','disassemblyTransport','transportAssembly','roomMove'];
    data.serviceKey=''; services.forEach(function(key){if(getFeeSelection_({serviceKey:key,regionKey:'capital',bedConfigKey:'single'}).serviceName===data.serviceName)data.serviceKey=key;});
    data.regionKey=data.regionName==='수도권'?'capital':'local';
    data.bedConfigKey=data.bedConfigName==='S+S'?'sPlusS':data.bedConfigName==='SS+SS'?'ssPlusSs':'single';
  } else data={name:raw[1],phone:raw[2],zipcode:raw[3],address:raw[4],detailAddress:raw[5],purchaseYear:raw[6],upgradeOption:raw[7],quantity:String(raw[8]),orderInfo:raw[9]||'',comments:raw[10]||'',email:raw[11]||''};
  var editable=status!=='취소' && status!=='완료' && status!=='발송 완료';
  if(MANAGEMENT_SERVICE==='showroom' && (data.date+' '+data.time)<Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm')) editable=false;
  var refund=null;
  if(c.paid && raw[c.paid-1]) {var account=managementText_(raw[c.account-1]).replace(/\D/g,'');refund={paid:raw[c.paid-1]==='입금 완료(고객 신고)',bank:raw[c.bank-1]||'',accountMasked:account?'*'.repeat(Math.max(0,account.length-4))+account.slice(-4):'',holder:raw[c.holder-1]||'',status:raw[c.refund-1]||''};}
  return {service:MANAGEMENT_SERVICE,rowId:n,status:status,editable:editable,data:data,refund:refund};
}

// 고객과 운영자 모두 같은 버튼·신청 전용 링크를 사용합니다. 텍스트 메일에도 전체 링크를 유지합니다.
function renderBirkitsEmail_(body) {
  function escape(value) {return String(value).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
  var content=String(body).split('\n').map(function(line) {
    var match=line.match(/^https:\/\/[^\s]+\/manage\.html\?action=(view|edit|cancel)#token=[A-Za-z0-9_.-]+$/);
    if(match) {
      var label={view:'신청 내용 확인',edit:'신청 내용 수정',cancel:'신청 취소'}[match[1]];
      return '<a href="'+escape(line)+'" style="display:inline-block;background:'+(match[1]==='cancel'?'#fff':'#252620')+';color:'+(match[1]==='cancel'?'#8a3024':'#fff')+';border:1px solid #aaa394;border-radius:8px;padding:12px 24px;margin:4px 0 12px;text-decoration:none;font-size:18px;font-weight:bold">'+label+'</a>';
    }
    var safe=escape(line);
    if(/^■ /.test(line))return '<strong style="font-size:19px">'+safe+'</strong>';
    return safe.replace(/https:\/\/[^\s<]+/g,function(url){return '<a href="'+url+'" style="color:#295a8f;word-break:break-all">'+url+'</a>';});
  }).join('<br>');
  return '<div style="font-family:Arial,sans-serif;max-width:620px;margin:auto;padding:24px;background:#fff;color:#161616;font-size:17px;line-height:1.8;word-break:keep-all"><div style="font-size:22px;letter-spacing:5px;border-bottom:1px solid #d9d3c9;padding-bottom:16px;margin-bottom:20px">birkits</div>'+content+'</div>';
}

function managementHandle_(request) {
  var sheet=managementSheet_(), lock=LockService.getScriptLock(), mutation=request.action!=='managedGet';
  if(mutation) lock.waitLock(10000);
  try {
    var owned=managementAuthorize_(sheet,request.token), n=owned.rowNumber, raw=owned.row;
    var current=managementPublic_(sheet,n,raw);
    if(request.action==='managedGet') return {result:'success',record:current};
    if(request.action==='managedCancel' && current.status==='취소') return {result:'success',record:current,alreadyCancelled:true,customerEmailSent:false};
    if(!current.editable) throw new Error('이미 처리되었거나 지난 일정으로 온라인 수정·취소가 불가능합니다. 담당자에게 문의해 주세요.');
    var mail, fresh, token;
    if(request.action==='managedUpdate') {
      var p=request.data||{}, previous=current.data;
      if(MANAGEMENT_SERVICE==='showroom') {
        fresh={name:previous.name,phone:previous.phone,email:managementText_(p.email).toLowerCase(),showroom:p.showroom,date:p.date,time:p.time,memo:p.memo||'',visitPromiseAccepted:p.visitPromiseAccepted};
        var error=validateReservationData(fresh); if(error) throw new Error(error);
        if(hasSlotConflict(sheet,fresh.showroom,fresh.date,fresh.time,n)) throw new Error('선택한 시간에 다른 예약이 있습니다. 다른 시간을 선택해 주세요.');
        sheet.getRange(n,4,1,4).setValues([[safeSheetText(fresh.showroom),safeSheetText(fresh.date+' '+fresh.time),safeSheetText(fresh.memo),safeSheetText(fresh.email)]]);
        sheet.getRange(n,11).setValue(new Date());
        if(reservationReminderKey(previous)!==reservationReminderKey(fresh))sheet.getRange(n,9,1,2).setValues([['','']]);
        token=managementIssue_(sheet,n,fresh.email);fresh._managementToken=token;
        mail={customerEmailSent:sendReservationNotifications('updated',fresh,previous)};
      } else if(MANAGEMENT_SERVICE==='moving') {
        fresh={}; ['depositTime','requestDate','isSecondHand','applicantName','applicantPhone','buyerName','buyerPhone','options','memo','addressBefore','addressAfter','addressBeforeBase','addressBeforeDetail','addressAfterBase','addressAfterDetail','customerEmail','serviceKey','regionKey','bedConfigKey'].forEach(function(key){fresh[key]=managementText_(p[key]);});
        fresh.customerEmail=normalizeCustomerEmail_(fresh.customerEmail);validateApplicationData_(fresh,true);
        var fee=formatFeeSummary_(getFeeSelection_(fresh));
        sheet.getRange(n,2,1,11).setValues([[safeCell_(fresh.depositTime),safeCell_(fresh.requestDate),safeCell_(fresh.isSecondHand),safeCell_(combineAddress_(fresh.addressBefore,fresh.addressAfter)),safeCell_(fresh.applicantName),safeCell_(fresh.applicantPhone),safeCell_(fresh.buyerName||'생략(동일)'),safeCell_(fresh.buyerPhone||'생략(동일)'),safeCell_(fresh.options),safeCell_(fresh.memo),fee]]);
        sheet.getRange(n,14,1,5).setValues([[new Date(),safeCell_(fresh.addressBeforeBase),safeCell_(fresh.addressBeforeDetail),safeCell_(fresh.addressAfterBase),safeCell_(fresh.addressAfterDetail)]]);
        sheet.getRange(n,23).setValue(safeCell_(fresh.customerEmail));
        token=managementIssue_(sheet,n,fresh.customerEmail);fresh._managementToken=token;
        mail=sendApplicationEmails_('update',fresh,combineAddress_(fresh.addressBefore,fresh.addressAfter),new Date());
      } else {
        fresh={};['name','phone','zipcode','address','detailAddress','purchaseYear','upgradeOption','quantity','orderInfo','comments','email'].forEach(function(key){fresh[key]=managementText_(p[key]);});
        fresh.email=normalizeEmail(fresh.email);
        if(!fresh.name || String(fresh.phone).replace(/\D/g,'').length<10 || !isValidEmail(fresh.email) || !fresh.address || !fresh.detailAddress || !/^[1-3]$/.test(fresh.quantity))throw new Error('필수 신청 정보를 정확히 입력해 주세요.');
        if(['2025년 이후','2024년 이전'].indexOf(fresh.purchaseYear)<0)throw new Error('구매 연도를 선택해 주세요.');
        var basePrice=fresh.purchaseYear==='2025년 이후'?15000:50000;
        var allowed=['자가교체('+basePrice.toLocaleString('ko-KR')+')','방문-수도권('+(basePrice+30000).toLocaleString('ko-KR')+')','방문-지방('+(basePrice+60000).toLocaleString('ko-KR')+')'];
        if(allowed.indexOf(fresh.upgradeOption)<0)throw new Error('업그레이드 방법을 다시 선택해 주세요.');
        fresh.phone=formatPhone(fresh.phone);fresh.zipcode=formatZip(fresh.zipcode);
        sheet.getRange(n,1,1,12).setValues([[new Date(),managementSafeCell_(fresh.name),managementSafeCell_(fresh.phone),managementSafeCell_(fresh.zipcode),managementSafeCell_(fresh.address),managementSafeCell_(fresh.detailAddress),fresh.purchaseYear,fresh.upgradeOption,fresh.quantity,managementSafeCell_(fresh.orderInfo),managementSafeCell_(fresh.comments),safeEmailCell(fresh.email)]]);
        token=managementIssue_(sheet,n,fresh.email);fresh._managementToken=token;
        mail=sendUpgradeApplicationEmails_('update',fresh,fresh.phone,fresh.zipcode,fresh.email);
      }
    } else if(request.action==='managedCancel') {
      if(request.confirmCancel!==true)throw new Error('취소 내용을 확인한 뒤 취소 확정 버튼을 눌러 주세요.');
      var c=managementConfig_(), refund=null;
      if(MANAGEMENT_SERVICE!=='showroom') {
        if(request.payment!=='paid' && request.payment!=='unpaid')throw new Error('입금 여부를 선택해 주세요.');
        refund={paid:request.payment==='paid',bank:'',account:'',holder:''};
        if(refund.paid) {
          refund.bank=managementText_(request.refundBank);refund.account=managementText_(request.refundAccount).replace(/[\s-]/g,'');refund.holder=managementText_(request.refundHolder);
          if(!refund.bank || refund.bank.length>40 || !/^\d{6,30}$/.test(refund.account) || !refund.holder || refund.holder.length>40 || /[\r\n]/.test(refund.bank+refund.holder))throw new Error('입금한 신청은 은행명·계좌번호·예금주를 모두 정확히 입력해야 취소할 수 있습니다.');
        }
        sheet.getRange(n,c.status,1,7).setValues([['취소',new Date(),refund.paid?'입금 완료(고객 신고)':'입금 전(고객 신고)',managementSafeCell_(refund.bank),refund.account,managementSafeCell_(refund.holder),refund.paid?'환불 대기':'입금 확인 필요']]);
      } else {sheet.getRange(n,4).setValue(CANCELLED_PREFIX+current.data.showroom);sheet.getRange(n,8).setValue('취소');}
      token=managementIssue_(sheet,n,managementEmail_(raw));fresh=current.data;fresh._managementToken=token;fresh._refund=refund;
      if(MANAGEMENT_SERVICE==='showroom')mail={customerEmailSent:sendReservationNotifications('cancelled',fresh)};
      else if(MANAGEMENT_SERVICE==='moving')mail=sendApplicationEmails_('cancel',fresh,'',new Date());
      else mail=sendUpgradeApplicationEmails_('cancel',fresh,fresh.phone,fresh.zipcode,fresh.email);
    } else throw new Error('지원하지 않는 관리 요청입니다.');
    SpreadsheetApp.flush();
    return {result:'success',record:managementPublic_(sheet,n,sheet.getRange(n,1,1,managementConfig_().width).getValues()[0]),token:token,customerEmailSent:Boolean(mail&&mail.customerEmailSent)};
  } finally {if(mutation)lock.releaseLock();}
}
