// ===================================================
// RTO Ninja — Google Apps Script
// Tracks office attendance & emails weekly reports
//
// HOW TO USE:
//   1. Paste this entire file into your Apps Script editor
//      (replacing whatever was there before)
//   2. Set Project Timezone to "America/Los_Angeles" in
//      Project Settings (gear icon)
//   3. Run setupTrigger() once to create the Sunday 9am trigger
//   4. Run testReport() to verify the email looks right
// ===================================================

// ===== CONFIG =====
var SHEET_ID              = "YOUR_GOOGLE_SHEET_ID";
var SHEET_TAB             = "Sheet1";
var VACATION_TAB          = "vacation";
var EMAIL_TO              = "YOUR_EMAIL_HERE";
var TARGET_DAYS           = 24;
var WEEKS_COUNTED         = 8;
var WEEKS_WINDOW          = 12;
var TIMEZONE              = "America/Los_Angeles";
var VACATION_LOOKAHEAD    = 35; // days — show forecast when trip starts within this window

// ===================================================
// TASKER ENDPOINT
// Tasker calls this URL when you arrive at the office.
// Logs today as an office day — weekdays only, no dupes.
// ===================================================
function doGet(e) {
  var now = new Date();
  var dayOfWeek = now.getDay(); // 0=Sun, 6=Sat

  // Block weekends
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return ContentService.createTextOutput("Weekend — not logged.");
  }

  var todayStr = Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd");
  var sheet = getSheet();

  if (isAlreadyLogged(sheet, todayStr)) {
    return ContentService.createTextOutput("Already logged: " + todayStr);
  }

  // Columns: Start Date | End Date | Vacation | Manual Entry
  sheet.appendRow([todayStr, todayStr, "", 1]);
  return ContentService.createTextOutput("Logged: " + todayStr);
}

// ===================================================
// SHEET HELPERS
// ===================================================
function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);
}

// Convert a cell value (Date object or "yyyy-MM-dd" string) to "yyyy-MM-dd"
function toDateStr(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE, "yyyy-MM-dd");
  }
  var s = String(val).trim();
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) return s;
  return "";
}

// Parse "yyyy-MM-dd" string to a UTC-noon Date (avoids DST shifts in arithmetic)
function parseDate(str) {
  if (!str || !str.match(/^\d{4}-\d{2}-\d{2}$/)) return null;
  var p = str.split("-");
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], 12, 0, 0));
}

// Check if dateStr (yyyy-MM-dd) already falls within any non-vacation row
function isAlreadyLogged(sheet, dateStr) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[2]).trim().toLowerCase() === "yes") continue; // vacation row
    var startStr = toDateStr(row[0]);
    var endStr   = row[1] ? toDateStr(row[1]) : startStr;
    if (!startStr) continue;
    if (!endStr)   endStr = startStr;
    // yyyy-MM-dd strings compare correctly lexicographically
    if (dateStr >= startStr && dateStr <= endStr) return true;
  }
  return false;
}

// ===================================================
// COMPLIANCE CALCULATION
// ===================================================

// Returns {dateStr: true} for every logged weekday office visit
function getOfficeDaysMap() {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  var map   = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[2]).trim().toLowerCase() === "yes") continue; // vacation
    if (!row[0]) continue;
    if (!row[3] || +row[3] === 0) continue; // skip if Manual Entry is blank or 0

    var startStr = toDateStr(row[0]);
    var endStr   = row[1] ? toDateStr(row[1]) : startStr;
    if (!startStr) continue;
    if (!endStr)   endStr = startStr;

    var cur = parseDate(startStr);
    var end = parseDate(endStr);
    if (!cur || !end) continue;

    while (cur <= end) {
      var dow = cur.getUTCDay();
      if (dow !== 0 && dow !== 6) { // weekdays only
        map[Utilities.formatDate(cur, TIMEZONE, "yyyy-MM-dd")] = true;
      }
      cur.setUTCDate(cur.getUTCDate() + 1);
    }
  }
  return map;
}

// Return the UTC-noon Date for the Monday of the ISO week containing `date`
function getWeekMonday(date) {
  var dow  = date.getUTCDay(); // 0=Sun
  var diff = (dow === 0) ? -6 : 1 - dow;
  var d    = new Date(date.getTime());
  d.setUTCDate(d.getUTCDate() + diff);
  return d;
}

// Count logged weekdays in the Mon–Fri week starting at weekMonday
function countDaysInWeek(officeDaysMap, weekMonday) {
  var count = 0;
  for (var d = 0; d < 5; d++) {
    var day = new Date(weekMonday.getTime());
    day.setUTCDate(weekMonday.getUTCDate() + d);
    if (officeDaysMap[Utilities.formatDate(day, TIMEZONE, "yyyy-MM-dd")]) count++;
  }
  return count;
}

// Main compliance calculation — runs as of today (Sunday)
function calculateCompliance() {
  var officeDaysMap = getOfficeDaysMap();

  // Parse today in PT
  var todayStr  = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd");
  var todayDate = parseDate(todayStr);
  var todayDow  = todayDate.getUTCDay(); // 0=Sun

  // On Sunday the Mon-Fri work week just finished, so include it in the window.
  // currentWeekMon = the Monday of the UPCOMING week (tomorrow on Sunday,
  // or next Monday on any other day). The 12-week window is the 12 weeks before it.
  var currentWeekMon;
  if (todayDow === 0) {
    // Sunday: next Monday is tomorrow
    currentWeekMon = new Date(todayDate.getTime());
    currentWeekMon.setUTCDate(todayDate.getUTCDate() + 1);
  } else {
    // Weekday/Saturday: skip to next Monday
    currentWeekMon = getWeekMonday(todayDate);
    currentWeekMon.setUTCDate(currentWeekMon.getUTCDate() + 7);
  }

  // nextMonday = currentWeekMon (the upcoming Mon we're recommending for)
  var nextMonday = new Date(currentWeekMon.getTime());
  var nextFriday = new Date(nextMonday.getTime());
  nextFriday.setUTCDate(nextMonday.getUTCDate() + 4);

  // Build the last WEEKS_WINDOW completed weeks (most recent first)
  // currentWeekMon is excluded ("current week never counted")
  var weeks = [];
  for (var w = 0; w < WEEKS_WINDOW; w++) {
    var wMon = new Date(currentWeekMon.getTime());
    wMon.setUTCDate(currentWeekMon.getUTCDate() - (w + 1) * 7);
    var wFri = new Date(wMon.getTime());
    wFri.setUTCDate(wMon.getUTCDate() + 4);
    weeks.push({
      weekNum : w + 1,
      monday  : wMon,
      friday  : wFri,
      count   : countDaysInWeek(officeDaysMap, wMon)
    });
  }

  // Best WEEKS_COUNTED of WEEKS_WINDOW
  var sorted   = weeks.slice().sort(function(a, b) { return b.count - a.count; });
  var bestSet  = {};
  for (var b = 0; b < WEEKS_COUNTED; b++) {
    if (sorted[b]) bestSet[sorted[b].weekNum] = true;
  }
  var totalBest = sorted.slice(0, WEEKS_COUNTED).reduce(function(s, w) { return s + w.count; }, 0);

  // Minimum days needed next week (projects to the Sunday after nextWeek completes)
  var daysNeeded = calculateDaysNeeded(officeDaysMap, nextMonday, TARGET_DAYS, WEEKS_COUNTED, WEEKS_WINDOW);

  // Vacation forecasts for trips starting within VACATION_LOOKAHEAD days
  var upcomingVacations   = getUpcomingVacations(nextMonday);
  var vacationForecasts   = upcomingVacations.map(function(vac) {
    return calculateVacationForecast(officeDaysMap, nextMonday, vac);
  });

  return {
    weeks            : weeks,
    bestSet          : bestSet,
    totalBest        : totalBest,
    target           : TARGET_DAYS,
    nextMonday       : nextMonday,
    nextFriday       : nextFriday,
    daysNeeded       : daysNeeded,
    vacationForecasts: vacationForecasts
  };
}

// Find the minimum days in nextWeek so compliance holds on the Sunday
// AFTER nextWeek completes.
//
// On that future Sunday, the new currentWeekMon = nextMonday + 7, and
// the 12-week window becomes:
//   w=0  nextWeek  (nextMonday)           ← we're solving for this
//   w=1  this week (nextMonday - 7)       ← already in the sheet
//   w=2            (nextMonday - 14)
//   ...
//   w=11 oldest    (nextMonday - 77)
function calculateDaysNeeded(officeDaysMap, nextMonday, target, weeksCount, weeksWindow) {
  // futureCurrentMon = the Monday of the week AFTER nextWeek completes
  var futureCurrentMon = new Date(nextMonday.getTime());
  futureCurrentMon.setUTCDate(nextMonday.getUTCDate() + 7);

  // Collect counts for weeks w=1..11 (w=0 is nextWeek, unknown)
  var knownCounts = [];
  for (var w = 1; w < weeksWindow; w++) {
    var wMon = new Date(futureCurrentMon.getTime());
    wMon.setUTCDate(futureCurrentMon.getUTCDate() - (w + 1) * 7);
    knownCounts.push(countDaysInWeek(officeDaysMap, wMon));
  }

  // Find minimum x (0–5) so that best weeksCount of [x, ...knownCounts] >= target
  for (var x = 0; x <= 5; x++) {
    var all    = [x].concat(knownCounts);
    var sorted = all.slice().sort(function(a, b) { return b - a; });
    var best   = sorted.slice(0, weeksCount).reduce(function(s, c) { return s + c; }, 0);
    if (best >= target) return x;
  }
  return 5; // Even 5 days may not be enough (very behind)
}

// ===================================================
// VACATION FORECAST
// ===================================================

// Read the "vacations" sheet and return trips that start within VACATION_LOOKAHEAD
// days of nextMonday and haven't fully ended yet.
function getUpcomingVacations(nextMonday) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(VACATION_TAB);
  if (!sheet) return [];

  var data    = sheet.getDataRange().getValues();
  var cutoff  = new Date(nextMonday.getTime());
  cutoff.setUTCDate(nextMonday.getUTCDate() + VACATION_LOOKAHEAD);
  var results = [];

  for (var i = 1; i < data.length; i++) { // row 0 = header
    var row = data[i];
    if (!row[0]) continue;
    var startStr = toDateStr(row[0]);
    var endStr   = row[1] ? toDateStr(row[1]) : startStr;
    var start    = parseDate(startStr);
    var end      = parseDate(endStr);
    if (!start || !end) continue;
    // Include if: trip hasn't fully ended AND starts within lookahead window
    if (end >= nextMonday && start <= cutoff) {
      results.push({ start: start, end: end, name: String(row[2] || "").trim() });
    }
  }
  results.sort(function(a, b) { return a.start - b.start; });
  return results;
}

// For a single vacation, find the minimum days/week needed in lead-up weeks
// so that compliance is still >= TARGET_DAYS on the Sunday after the trip ends.
function calculateVacationForecast(officeDaysMap, nextMonday, vacation) {
  // The next compliance report after the trip runs on the Sunday after vacation ends.
  // With our Sunday logic, that Sunday's currentWeekMon = Sunday + 1 (next Monday).
  var endDow         = vacation.end.getUTCDay();
  var daysToSunday   = (7 - endDow) % 7; // 0 if already Sunday
  var sundayAfterVac = new Date(vacation.end.getTime());
  sundayAfterVac.setUTCDate(vacation.end.getUTCDate() + daysToSunday);
  var futureCurrentMon = new Date(sundayAfterVac.getTime());
  futureCurrentMon.setUTCDate(sundayAfterVac.getUTCDate() + 1);

  // Build the 12-week compliance window for that future Sunday
  var windowWeeks = [];
  for (var w = 0; w < WEEKS_WINDOW; w++) {
    var wMon = new Date(futureCurrentMon.getTime());
    wMon.setUTCDate(futureCurrentMon.getUTCDate() - (w + 1) * 7);
    var wFri = new Date(wMon.getTime());
    wFri.setUTCDate(wMon.getUTCDate() + 4);

    // Vacation week = Mon–Fri overlaps with vacation dates
    var isVac    = wMon <= vacation.end && wFri >= vacation.start;
    // Lead-up week = on/after nextMonday, not a vacation week, not in the past
    var isLeadUp = !isVac && wMon >= nextMonday;

    windowWeeks.push({
      monday  : wMon,
      friday  : wFri,
      isVac   : isVac,
      isLeadUp: isLeadUp,
      known   : (!isVac && !isLeadUp) ? countDaysInWeek(officeDaysMap, wMon) : null
    });
  }

  var leadUpWeeks = windowWeeks.filter(function(w) { return w.isLeadUp; });

  // Find minimum uniform days/week in lead-up weeks so best WEEKS_COUNTED of 12 >= TARGET
  for (var x = 0; x <= 5; x++) {
    var counts = windowWeeks.map(function(w) {
      if (w.isVac)    return 0;
      if (w.isLeadUp) return x;
      return w.known;
    });
    var sorted = counts.slice().sort(function(a, b) { return b - a; });
    var best   = sorted.slice(0, WEEKS_COUNTED).reduce(function(s, c) { return s + c; }, 0);
    if (best >= TARGET_DAYS) {
      return {
        possible       : true,
        daysPerWeek    : x,
        numLeadUpWeeks : leadUpWeeks.length,
        leadUpWeeks    : leadUpWeeks,
        vacation       : vacation,
        sundayAfterVac : sundayAfterVac
      };
    }
  }

  // Even 5 days/week every lead-up week isn't enough
  return {
    possible       : false,
    numLeadUpWeeks : leadUpWeeks.length,
    leadUpWeeks    : leadUpWeeks,
    vacation       : vacation,
    sundayAfterVac : sundayAfterVac
  };
}

// Render one vacation forecast as an HTML block for the email
function buildVacationForecastSection(forecast) {
  var vac  = forecast.vacation;
  var html = "<div style='border-left:4px solid #f9a825;padding:10px 14px;"
           + "margin-bottom:12px;background:#fffde7;'>";
  var tripLabel = vac.name ? vac.name + " (" + fmtDate(vac.start) + " – " + fmtDate(vac.end) + ")"
                           : fmtDate(vac.start) + " – " + fmtDate(vac.end);
  html += "<strong>✈️ " + tripLabel + "</strong>";

  if (forecast.numLeadUpWeeks === 0) {
    html += "<br><span style='color:#c62828;'>⚠️ No lead-up weeks available — "
          + "compliance will slip during this trip.</span>";
  } else if (!forecast.possible) {
    html += "<br><span style='color:#c62828;'>⚠️ Even going in 5 days/week before your trip, "
          + "compliance cannot be maintained through this vacation.</span>";
  } else if (forecast.daysPerWeek === 0) {
    html += "<br>✅ You're banked enough — no extra effort needed for this trip.";
  } else {
    html += "<p style='margin:6px 0;'>Aim for at least <strong>" + forecast.daysPerWeek
          + " day" + (forecast.daysPerWeek !== 1 ? "s" : "") + " per week</strong> in the "
          + forecast.numLeadUpWeeks + " lead-up week"
          + (forecast.numLeadUpWeeks !== 1 ? "s" : "") + " before you leave:</p>";

    html += "<table border='1' cellpadding='4' cellspacing='0' "
          + "style='border-collapse:collapse;font-family:monospace;font-size:12px;'>";
    html += "<tr style='background:#f1f3f4;'><th>Week</th><th>Target days</th></tr>";
    forecast.leadUpWeeks.forEach(function(w) {
      html += "<tr><td>" + fmtDate(w.monday) + " – " + fmtDate(w.friday) + "</td>"
            + "<td style='text-align:center;'>" + forecast.daysPerWeek + "</td></tr>";
    });
    html += "</table>";
  }

  html += "</div>";
  return html;
}

// ===================================================
// EMAIL
// ===================================================
function fmtDate(date) {
  return Utilities.formatDate(date, TIMEZONE, "MMM d");
}

function buildEmail(result) {
  var weekLabel = fmtDate(result.nextMonday) + "–" + fmtDate(result.nextFriday);
  var pct       = Math.round((result.totalBest / result.target) * 100);
  var onTrack   = result.totalBest >= result.target;
  var statusIcon = onTrack ? "✅" : "⚠️";

  var rec;
  if (result.daysNeeded === 0) {
    rec = "You're ahead — no minimum required. Every day is a bonus! 🎉";
  } else if (result.daysNeeded <= 5) {
    rec = "Go in at least <strong>" + result.daysNeeded + " day"
        + (result.daysNeeded !== 1 ? "s" : "") + "</strong> to stay on track.";
  } else {
    rec = "⚠️ Even 5 days won't fully close the gap this week — keep chipping away.";
  }

  var html = "";
  html += "<h2 style='color:#1a73e8;margin-bottom:4px;'>RTO Ninja</h2>";
  html += "<p style='color:#555;margin-top:0;'>Weekly compliance report — sent every Sunday at 9am PT</p>";
  html += "<hr>";

  html += "<table style='border-collapse:collapse;margin:12px 0;'>";
  html += "<tr><td style='padding:4px 16px 4px 0;font-weight:bold;'>Compliance (best "
        + WEEKS_COUNTED + " of " + WEEKS_WINDOW + " weeks):</td>";
  html += "<td>" + result.totalBest + " / " + result.target + " days (" + pct + "%) " + statusIcon + "</td></tr>";
  html += "<tr><td style='padding:4px 16px 4px 0;font-weight:bold;vertical-align:top;'>"
        + "Next week's recommendation<br>(" + weekLabel + "):</td>";
  html += "<td style='vertical-align:top;'>" + rec + "</td></tr>";
  html += "</table>";

  html += "<h3 style='margin-bottom:6px;'>Last 12 completed weeks</h3>";
  html += "<table border='1' cellpadding='6' cellspacing='0' "
        + "style='border-collapse:collapse;font-family:monospace;font-size:13px;'>";
  html += "<tr style='background:#f1f3f4;text-align:left;'>"
        + "<th>#</th><th>Mon</th><th>Fri</th><th>Days</th><th></th></tr>";

  for (var i = 0; i < result.weeks.length; i++) {
    var w      = result.weeks[i];
    var inBest = result.bestSet[w.weekNum];
    var bg     = inBest ? "#e8f5e9" : "#fff";
    var star   = inBest ? " ⭐" : "";
    var note   = inBest ? "<span style='color:#2e7d32;font-size:11px;'>counted</span>" : "";
    html += "<tr style='background:" + bg + ";'>";
    html += "<td style='text-align:center;'>" + w.weekNum + "</td>";
    html += "<td>" + fmtDate(w.monday) + "</td>";
    html += "<td>" + fmtDate(w.friday) + "</td>";
    html += "<td style='text-align:center;'>" + w.count + star + "</td>";
    html += "<td>" + note + "</td>";
    html += "</tr>";
  }
  html += "</table>";

  html += "<p style='color:#888;font-size:11px;margin-top:12px;'>"
        + "⭐ = counted in best " + WEEKS_COUNTED + " &nbsp;|&nbsp; "
        + "Target: " + result.target + " days from best " + WEEKS_COUNTED
        + " of " + WEEKS_WINDOW + " completed weeks</p>";

  // Vacation forecasts
  if (result.vacationForecasts && result.vacationForecasts.length > 0) {
    html += "<h3 style='margin-bottom:6px;'>✈️ Upcoming Trip Forecast</h3>";
    result.vacationForecasts.forEach(function(forecast) {
      html += buildVacationForecastSection(forecast);
    });
  }

  return html;
}

// ===================================================
// MAIN: SEND WEEKLY REPORT
// ===================================================
function sendWeeklyReport() {
  var result    = calculateCompliance();
  var htmlBody  = buildEmail(result);
  var weekLabel = fmtDate(result.nextMonday) + "–" + fmtDate(result.nextFriday);
  var subject   = "RTO Ninja — Next week (" + weekLabel + "): "
                + result.daysNeeded + " day" + (result.daysNeeded !== 1 ? "s" : "")
                + " needed  |  " + result.totalBest + "/" + result.target + " current";

  GmailApp.sendEmail(EMAIL_TO, subject, "Please view this email in HTML.", { htmlBody: htmlBody });
  Logger.log("Report sent to " + EMAIL_TO);
}

// ===================================================
// SETUP — run once after pasting this script
// ===================================================
function setupTrigger() {
  // Remove any existing sendWeeklyReport triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "sendWeeklyReport") ScriptApp.deleteTrigger(t);
  });

  // Sunday at 9am (uses the timezone set in Apps Script Project Settings)
  ScriptApp.newTrigger("sendWeeklyReport")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(9)
    .create();

  Logger.log("✅ Trigger set: sendWeeklyReport every Sunday ~9am");
}

// ===================================================
// TEST — run manually to send a test email right now
// ===================================================
function testReport() {
  var result = calculateCompliance();
  Logger.log("=== Compliance Summary ===");
  Logger.log("Total best " + WEEKS_COUNTED + ": " + result.totalBest + " / " + result.target);
  Logger.log("Days needed next week: " + result.daysNeeded);
  Logger.log("Next week: " + fmtDate(result.nextMonday) + " – " + fmtDate(result.nextFriday));
  Logger.log("Sending test email...");
  sendWeeklyReport();
  Logger.log("Done! Check " + EMAIL_TO);
}
