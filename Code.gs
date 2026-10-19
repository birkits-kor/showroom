var NOTIFY_EMAIL = PropertiesService.getScriptProperties().getProperty("NOTIFY_EMAIL") || "";
var CANCELLED_PREFIX = "[취소] ";
var SERVICE_VERSION = "2026-10-19-no-show-v1";
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

    var manageUrl = resolveReservationManageUrl(data.pageUrl);
    sendNotification(
      "[버킷츠] 새로운 쇼룸 예약이 접수되었습니다!",
      "버킷츠 쇼룸에 새로운 방문 예약이 접수되었습니다.\n\n" +
      "👤 고객명: " + data.name + "\n" +
      "📞 연락처: " + data.phone + "\n" +
      "🏢 방문지점: " + data.showroom + "\n" +
      "⏰ 방문일정: " + data.date + " " + data.time + "\n" +
      "📝 남긴말씀: " + (data.memo ? data.memo : "없음") + "\n\n" +
      "🔗 예약 확인·수정·취소 링크:\n" +
      (manageUrl || "예약 페이지 주소를 확인해 주세요.") + "\n"
    );

    var customerEmailSent = sendCustomerNotification(
      normalizeEmail(data.email),
      "[버킷츠] 쇼룸 방문 예약이 접수되었습니다.",
      buildCustomerMessage("created", data)
    );

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

    sendNotification(
      "[버킷츠] 쇼룸 예약이 변경되었습니다.",
      "쇼룸 방문 예약이 온라인으로 변경되었습니다.\n\n" +
      "👤 고객명: " + reservation.name + "\n" +
      "📞 연락처: " + reservation.phone + "\n" +
      "기존 일정: " + reservation.showroom + " / " + reservation.date + " " + reservation.time + "\n" +
      "변경 일정: " + data.showroom + " / " + data.date + " " + data.time + "\n" +
      "📝 남긴말씀: " + (data.memo ? data.memo : "없음") + "\n\n" +
      "🔗 예약 확인·수정·취소 링크:\n" + resolveReservationManageUrl(data.pageUrl) + "\n"
    );

    var customerEmailSent = sendCustomerNotification(
      normalizeEmail(data.email),
      "[버킷츠] 쇼룸 방문 예약이 변경되었습니다.",
      buildCustomerMessage("updated", {
        name: reservation.name, showroom: data.showroom, date: data.date,
        time: data.time, memo: data.memo
      }, reservation)
    );

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

    sendNotification(
      "[버킷츠] 쇼룸 예약이 취소되었습니다.",
      "쇼룸 방문 예약이 온라인으로 취소되었습니다.\n\n" +
      "👤 고객명: " + reservation.name + "\n" +
      "📞 연락처: " + reservation.phone + "\n" +
      "🏢 방문지점: " + reservation.showroom + "\n" +
      "⏰ 기존 일정: " + reservation.date + " " + reservation.time + "\n\n" +
      "🔗 예약 확인·수정·취소 링크:\n" + resolveReservationManageUrl(data.pageUrl) + "\n"
    );

    var customerEmailSent = reservation.email ? sendCustomerNotification(
      reservation.email,
      "[버킷츠] 쇼룸 방문 예약이 취소되었습니다.",
      buildCustomerMessage("cancelled", reservation)
    ) : false;

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
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (error) {
    console.error("이메일 발송 실패: " + error.message);
  }
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
      name: "버킷츠 쇼룸"
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

function reservationMailLinks() {
  var base = resolveReservationManageUrl("").split("#")[0];
  return {
    view: base + "#manageReservation",
    change: base + "?manageAction=change#manageReservation",
    cancel: base + "?manageAction=cancel#manageReservation"
  };
}

function buildCustomerMessage(event, reservation, previous) {
  var links = reservationMailLinks();
  var title = event === "reminder" ? "내일은 버킷츠 쇼룸 방문 예약일입니다." :
    event === "updated" ? "버킷츠 쇼룸 방문 예약이 변경되었습니다." :
    event === "cancelled" ? "버킷츠 쇼룸 방문 예약이 취소되었습니다." : "버킷츠 쇼룸 방문 예약이 접수되었습니다.";
  var body = "안녕하세요, " + cleanText(reservation.name) + "님.\n" + title + "\n\n";
  if (previous) body += "기존 일정: " + previous.showroom + " / " + previous.date + " " + previous.time + "\n\n";
  body += "🏢 방문지점: " + reservation.showroom + "\n" +
    "⏰ " + (event === "cancelled" ? "취소 일정" : "방문일정") + ": " + reservation.date + " " + reservation.time + "\n";
  if (event !== "cancelled") body += "상담 및 제품 체험: 약 1시간\n";
  body += "📝 남긴말씀: " + (reservation.memo ? reservation.memo : "없음") + "\n\n";
  var mapUrl = SHOWROOM_MAP_URLS[reservation.showroom];
  if (mapUrl) body += "📍 매장 전화번호 확인·길찾기(네이버 지도):\n" + mapUrl + "\n\n";
  body += "🔗 예약 확인·수정·취소 링크:\n" + links.view + "\n\n" +
    "방문 일정 변경:\n" + links.change + "\n\n" +
    "예약 취소:\n" + links.cancel + "\n\n" +
    "예약할 때 입력하신 성함과 연락처로 조회하실 수 있습니다.\n";
  if (event === "cancelled") {
    body += "\n다음에 편하신 일정으로 다시 찾아 주세요.\n";
  } else {
    body += "\n담당자가 고객님의 예약 시간을 비워두고 체험을 준비합니다.\n" +
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
  var links = reservationMailLinks();
  var text = escapeEmailHtml(body).replace(/https:\/\/[^\s<]+/g, function(url) {
    return '<a href="' + url + '" style="color:#295a8f;word-break:break-all">' + url + '</a>';
  }).replace(/\n/g, "<br>");
  return '<div style="font-family:Arial,sans-serif;max-width:620px;margin:auto;padding:24px;color:#161616;font-size:17px;line-height:1.7">' +
    '<h1 style="font-size:24px">버킷츠 쇼룸 예약 안내</h1>' + text +
    '<p style="margin-top:24px"><a style="display:block;padding:15px;margin-bottom:10px;background:#161616;color:#fff;text-align:center;text-decoration:none;border-radius:9px" href="' + links.view + '">예약 내용 확인</a>' +
    '<a style="display:block;padding:15px;margin-bottom:10px;background:#ede8df;color:#161616;text-align:center;text-decoration:none;border-radius:9px" href="' + links.change + '">방문 일정 변경</a>' +
    '<a style="display:block;padding:15px;border:1px solid #b74337;color:#a5352b;text-align:center;text-decoration:none;border-radius:9px" href="' + links.cancel + '">예약 취소</a></p></div>';
}

// 설치 시 시트와 오전 9시경 트리거만 준비합니다. 고객 메일은 발송하지 않습니다.
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
    var triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
      return trigger.getHandlerFunction() === "sendTomorrowReservationReminders";
    });
    if (!triggers.length) {
      ScriptApp.newTrigger("sendTomorrowReservationReminders").timeBased()
        .atHour(9).nearMinute(0).everyDays(1).inTimezone("Asia/Seoul").create();
    }
    SpreadsheetApp.flush();
    console.log(JSON.stringify({
      version: SERVICE_VERSION, setup: "complete", reminderTime: "매일 오전 9시경 (Asia/Seoul)",
      reminderTriggers: Math.max(1, triggers.length), reminderEmailsSentDuringSetup: 0
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
        name: cleanText(row[1]), showroom: cleanText(row[3]), date: parsed.date,
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
  var triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === "sendTomorrowReservationReminders";
  });
  var result = {
    version: SERVICE_VERSION,
    spreadsheetId: sheet.getParent().getId(),
    sheet: sheet.getName(),
    headers: sheet.getRange(1, 7, 1, 6).getValues()[0],
    reminderTriggers: triggers.length,
    lastReminderRun: PropertiesService.getScriptProperties().getProperty("LAST_REMINDER_RUN") || "아직 실행되지 않음"
  };
  console.log(JSON.stringify(result));
  return result;
}
