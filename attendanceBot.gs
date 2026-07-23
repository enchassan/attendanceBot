// =================================================================
// SPRINT 1: CORE ROUTER & PARSER
// =================================================================

function doPost(e) {
  try {
    if (e.parameter.payload) {
      // Interactive payload handling (for future button clicks/approvals)
      return ContentService.createTextOutput(""); 
    }
    return handleCommandRouter(e.parameter);
  } catch (error) {
    return sendEphemeralResponse("❌ *System Error:* " + error.toString());
  }
}

function handleCommandRouter(params) {
  const command = params.command;
  const rawText = params.text ? params.text.trim() : "";
  const args = parseArgs(rawText);

  switch (command) {
    case "/login":
      return handleLogin(params, args);
    case "/logout":
      return handleLogout(params, args);
    case "/attendance":
      return handleAttendance(params, args);
    case "/leave-req":
      return handleLeave(params, args);
    case "/help":
      return handleHelp(params, args);
    default:
      return sendEphemeralResponse("⚠️ Unknown command. Type `/help` for available commands.");
  }
}

// =================================================================
// ARGUMENT PARSER
// Extracts flags like --reason : Traffic, --wfh, --date : -1
// =================================================================
function parseArgs(rawText) {
  const args = {};
  if (!rawText) return args;
  
  // Split by "--" and clean up
  const parts = rawText.split('--').map(p => p.trim()).filter(Boolean);
  
  parts.forEach(part => {
    if (part === "[=]") {
      args["today_flag"] = true;
      return;
    }

    // Match the key and the optional value separated by space or colon
    const match = part.match(/^([^\s:]+)(?:[\s:]+(.*))?$/);
    if (match) {
      const key = match[1].toLowerCase();
      let value = match[2] ? match[2].trim() : true; // Default to true for boolean flags like --wfh
      
      // Strip quotes if user typed --reason "Stuck in traffic"
      if (typeof value === "string") {
        value = value.replace(/^["']|["']$/g, '');
      }
      args[key] = value;
    }
  });
  return args;
}

// =================================================================
// DYNAMIC DATE & TIME RESOLVERS
// Handles '=', '-n', '+n' based on GMT+5 (Lahore)
// =================================================================
function resolveDate(dateInput) {
  const d = new Date();
  if (!dateInput || dateInput === "=" || dateInput === true) {
    return Utilities.formatDate(d, "GMT+5", "yyyy-MM-dd");
  }
  
  if (dateInput.startsWith("-") || dateInput.startsWith("+")) {
    const days = parseInt(dateInput, 10);
    d.setDate(d.getDate() + days);
    return Utilities.formatDate(d, "GMT+5", "yyyy-MM-dd");
  }
  
  // Assume exact format YYYY-MM-DD
  return dateInput;
}

function resolveTime(timeInput) {
  const d = new Date();
  if (!timeInput || timeInput === "=" || timeInput === true) {
    return Utilities.formatDate(d, "GMT+5", "HH:mm");
  }
  // Assume exact format HH:MM
  return timeInput;
}

function resolveMonth(monthInput) {
  const d = new Date();
  // Apps Script getMonth() is 0-indexed, so we add 1
  const currentMonth = parseInt(Utilities.formatDate(d, "GMT+5", "MM"), 10);
  
  if (!monthInput || monthInput === "=" || monthInput === true) return currentMonth;
  
  if (monthInput.startsWith("-") || monthInput.startsWith("+")) {
    return currentMonth + parseInt(monthInput, 10);
  }
  
  return parseInt(monthInput, 10);
}

// =================================================================
// UTILS
// =================================================================
function sendEphemeralResponse(text) {
  return ContentService.createTextOutput(JSON.stringify({ response_type: "ephemeral", text: text }))
    .setMimeType(ContentService.MimeType.JSON);
}

// =================================================================
// SPRINT 2: /LOGIN & /LOGOUT WITH SPREADSHEET INTEGRATION
// =================================================================

function handleLogin(params, args) {
  const userId = params.user_id;
  const userName = getSlackRealName(userId) || params.user_name;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  
  const targetDate = resolveDate(args.date);
  const targetTime = resolveTime(args.time);
  
  const isReplaced = (args.replace === "true" || args.replace === true);
  const isWfh = (args.wfh === "true" || args.wfh === true);
  const reason = args.reason || "";
  
  const actualTimestamp = Utilities.formatDate(new Date(), "GMT+5", "yyyy-MM-dd HH:mm:ss");
  const todayStr = actualTimestamp.split(" ")[0]; // Extracts just the YYYY-MM-DD part
  
  // =================================================================
  // EXCEPTION FIX: PRIOR DATE STRICT REPLACEMENT
  // =================================================================
  if (targetDate < todayStr && !isReplaced) {
    return sendEphemeralResponse(`⚠️ *Missing Replace Flag!* You are trying to log attendance for a past date (*${targetDate}*). You must use the \`--replace true\` flag to authorize logging prior attendance.\nExample: \`/login --date ${targetDate} --replace true --time HH:MM\``);
  }

  const data = sheet.getDataRange().getDisplayValues();
  let foundRow = -1;

  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = String(data[i][1]).trim();
    const rowSlackId = String(data[i][15]).trim(); // Column P
    if (rowDate === targetDate && rowSlackId === userId) {
      foundRow = i + 1; 
      break;
    }
  }

  if (foundRow !== -1) {
    const existingLoginTime = String(data[foundRow - 1][3]).trim(); // Column D
    const existingLogoutTime = String(data[foundRow - 1][8]).trim(); // Column I
    
    // Prevent duplicate logins without --replace
    if (existingLoginTime !== "" && !isReplaced) {
      return sendEphemeralResponse(`⚠️ *Already Logged In!* You checked in at *${existingLoginTime}*. If you are trying to overwrite this, you must use the \`--replace true\` flag.`);
    }

    // VALIDATION FIX: Prevent login time > logout time
    if (existingLogoutTime !== "") {
      const loginDec = timeToDec(targetTime);
      const logoutDec = timeToDec(existingLogoutTime);
      
      if (loginDec > logoutDec) {
        return sendEphemeralResponse(`⚠️ *Invalid Time!* Your login time (*${targetTime}*) cannot be later than your existing logout time (*${existingLogoutTime}*). Please use \`/login --replace true --time HH:MM\`.`);
      }
    }

    // Process Overwrite / Replace for Login Data
    sheet.getRange(foundRow, 4).setValue(`'${targetTime}`);       // D: Login Time
    sheet.getRange(foundRow, 5).setValue(isReplaced);             // E: Login Replaced?
    sheet.getRange(foundRow, 6).setValue(actualTimestamp);        // F: Login Actual Action
    sheet.getRange(foundRow, 7).setValue(reason);                 // G: Late Login Reason
    sheet.getRange(foundRow, 8).setValue(isWfh);                  // H: Work From Home?
    
    // RECALCULATION FIX: If they had a "Missing In" record, calculate the hours now
    if (existingLogoutTime !== "") {
      const diffHours = calcDecimalHours24(targetTime, existingLogoutTime);
      const totalHoursFormatted = Number(diffHours.toFixed(2));
      const status = totalHoursFormatted >= 8.5 ? "Full Day 🟢" : "Half Day 🟡";

      sheet.getRange(foundRow, 13).setValue(totalHoursFormatted); // M: Total Hours
      sheet.getRange(foundRow, 14).setValue(status);              // N: Status

      return sendEphemeralResponse(`⚠️ *Attendance Recovered:* Your login for *${targetDate}* is set to *${targetTime}*. Since you already logged out at *${existingLogoutTime}*, your total hours are now *${totalHoursFormatted} hrs* (${status}).`);
    }

    return sendEphemeralResponse(`⚠️ *Attendance Overwritten:* Your login for *${targetDate}* is now updated to *${targetTime}*.\n*Note:* This overwrite has been flagged for HR review.`);
  }

  // Create a brand new record
  sheet.appendRow([
    "=ROW()-1", targetDate, userName, `'${targetTime}`, isReplaced, 
    actualTimestamp, reason, isWfh, "", "", "", "", "", "In Progress ⏳", "", userId
  ]);

  const wfhText = isWfh ? " 🏠 *(Working From Home)*" : "";
  return sendEphemeralResponse(`✅ Successfully *Logged In 🟢* at *${targetTime}* for *${targetDate}*${wfhText}.\n📝 *Reason:* ${reason || "None"}`);
}

// Small helper function specifically for the time validation check
function timeToDec(t) {
  const parts = String(t).split(":");
  if (parts.length !== 2) return 0;
  return parseInt(parts[0], 10) + (parseInt(parts[1], 10) / 60);
}

function handleLogout(params, args) {
  const userId = params.user_id;
  const userName = getSlackRealName(userId) || params.user_name;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  
  const targetDate = resolveDate(args.date);
  const targetTime = resolveTime(args.time);
  
  const isReplaced = (args.replace === "true" || args.replace === true);
  const reason = args.reason || "";
  const actualTimestamp = Utilities.formatDate(new Date(), "GMT+5", "yyyy-MM-dd HH:mm:ss");
  const todayStr = actualTimestamp.split(" ")[0]; 
  
  // =================================================================
  // EXCEPTION FIX: PRIOR DATE STRICT REPLACEMENT
  // =================================================================
  if (targetDate < todayStr && !isReplaced) {
    return sendEphemeralResponse(`⚠️ *Missing Replace Flag!* You are trying to log out for a past date (*${targetDate}*). You must use the \`--replace true\` flag to authorize modifying prior attendance.\nExample: \`/logout --date ${targetDate} --replace true --time HH:MM\``);
  }

  const data = sheet.getDataRange().getDisplayValues();
  let foundRow = -1;
  let loginTime = "";

  for (let i = data.length - 1; i >= 1; i--) {
    const rowDate = String(data[i][1]).trim();
    const rowSlackId = String(data[i][15]).trim(); // Column P
    if (rowDate === targetDate && rowSlackId === userId) {
      foundRow = i + 1;
      loginTime = String(data[i][3]).trim(); // Column D
      break;
    }
  }

  if (foundRow === -1 || !loginTime) {
    sheet.appendRow([
      "=ROW()-1", targetDate, userName, "", false, "", "", false,
      `'${targetTime}`, isReplaced, actualTimestamp, reason, 0, "Missing In 🔴", "", userId
    ]);
    return sendEphemeralResponse(`⚠️ *Warning:* You logged out at *${targetTime}*, but I couldn't find your Login for today. An incomplete record has been generated and flagged for HR.`);
  }

  const existingLogoutTime = String(data[foundRow - 1][8]).trim(); // Column I

  if (existingLogoutTime !== "" && !isReplaced) {
    return sendEphemeralResponse(`⚠️ *Already Logged Out!* You checked out at *${existingLogoutTime}*. To overwrite, use the \`--replace true\` flag.`);
  }

  // EXCEPTION FIX: Prevent logout time < login time on the same date
  const loginDec = timeToDec(loginTime);
  const targetLogoutDec = timeToDec(targetTime);

  if (targetLogoutDec < loginDec) {
    return sendEphemeralResponse(`⚠️ *Invalid Time!* Your logout time (*${targetTime}*) cannot be earlier than your login time (*${loginTime}*). Please use \`/logout --replace true --time HH:MM\`.`);
  }

  const diffHours = calcDecimalHours24(loginTime, targetTime);
  const totalHoursFormatted = Number(diffHours.toFixed(2));
  const status = totalHoursFormatted >= 8.5 ? "Full Day 🟢" : "Half Day 🟡";

  sheet.getRange(foundRow, 9).setValue(`'${targetTime}`);          
  sheet.getRange(foundRow, 10).setValue(isReplaced);               
  sheet.getRange(foundRow, 11).setValue(actualTimestamp);          
  sheet.getRange(foundRow, 12).setValue(reason);                   
  sheet.getRange(foundRow, 13).setValue(totalHoursFormatted);      
  sheet.getRange(foundRow, 14).setValue(status);                   

  if (isReplaced) {
    return sendEphemeralResponse(`⚠️ *Logout Overwritten:* Your logout for *${targetDate}* is now *${targetTime}*. Total logged: ${totalHoursFormatted} hrs.\n*Note:* Overwrite flagged for HR review.`);
  }

  return sendEphemeralResponse(`✅ Successfully *Logged Out 🔴* at *${targetTime}*.\n⏱️ *Hours logged:* ${totalHoursFormatted} hrs (${status})\n📝 *Reason:* ${reason || "None"}`);
}

// =================================================================
// HELPER FUNCTIONS 
// =================================================================

function calcDecimalHours24(timeInStr, timeOutStr) {
  function toDec(t) {
    const parts = String(t).split(":");
    if (parts.length !== 2) return 0;
    const h = parseInt(parts[0], 10);
    const min = parseInt(parts[1], 10);
    return h + (min / 60);
  }
  
  let diff = toDec(timeOutStr) - toDec(timeInStr);
  if (diff < 0) diff += 24; // Handles working past midnight
  return diff;
}

function getSlackRealName(userId) {
  try {
    const response = UrlFetchApp.fetch(`https://slack.com/api/users.info?user=${userId}`, {
      method: "get",
      headers: { Authorization: "Bearer " + SLACK_BOT_TOKEN } // Ensure you have this token defined in your script properties
    });
    const json = JSON.parse(response.getContentText());
    if (json.ok && json.user && json.user.profile) {
      return json.user.profile.display_name || json.user.profile.real_name || json.user.name;
    }
  } catch (error) {}
  return null;
}

// =================================================================
// SPRINT 3: /ATTENDANCE COMMAND (RETRIEVAL & VALIDATION)
// =================================================================

function handleAttendance(params, args) {
  const userId = params.user_id;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");
  
  // Get current date constraints
  const now = new Date();
  const currentYear = parseInt(Utilities.formatDate(now, "GMT+5", "yyyy"), 10);
  const currentMonth = parseInt(Utilities.formatDate(now, "GMT+5", "MM"), 10);
  const todayString = Utilities.formatDate(now, "GMT+5", "yyyy-MM-dd");

  const isTodayMode = args.today_flag === true;
  let targetMonth = null;

  // 1. Evaluate user intent and validate constraints
  if (isTodayMode) {
    // Valid request for today
  } else if (args.month) {
    targetMonth = resolveMonth(args.month); // resolveMonth handles "=", "-n", "+n", or specific digits
    
    // Strict restriction: Cannot request future months, which implies asking for previous year's data
    if (targetMonth > currentMonth) {
      return sendEphemeralResponse(`⚠️ *Access Restricted:* You requested data for month ${targetMonth}, but we are currently in month ${currentMonth}.\n\nIf you want to check your past year's attendance, you have to request the HR Executive.`);
    }
    
    if (targetMonth < 1 || targetMonth > 12) {
      return sendEphemeralResponse("⚠️ *Invalid Month:* Please provide a valid month between 1 and 12, or use `=` for the current month.");
    }
  } else {
    return sendEphemeralResponse("⚠️ *Missing Parameters:* Please specify what you want to check.\nUse `/attendance --[=]` for today, or `/attendance --month =` for this month.");
  }

  // 2. Fetch the data
  const data = sheet.getDataRange().getDisplayValues();
  let results = [];

  // We loop forward for monthly reports so the dates are in chronological order
  for (let i = 1; i < data.length; i++) {
    const rowSlackId = String(data[i][15]).trim(); // Column P
    const rowDateStr = String(data[i][1]).trim();  // Column B (YYYY-MM-DD)
    
    if (rowSlackId !== userId || !rowDateStr) continue;

    if (isTodayMode) {
      if (rowDateStr === todayString) {
        results.push(data[i]);
      }
    } else if (targetMonth !== null) {
      const parts = rowDateStr.split('-');
      if (parts.length === 3) {
        const rowYear = parseInt(parts[0], 10);
        const rowMonth = parseInt(parts[1], 10);
        
        if (rowYear === currentYear && rowMonth === targetMonth) {
          results.push(data[i]);
        }
      }
    }
  }

  // 3. Format the Slack Output
  if (results.length === 0) {
    if (isTodayMode) {
      return sendEphemeralResponse(`📭 No attendance record found for today (*${todayString}*).`);
    } else {
      return sendEphemeralResponse(`📭 No attendance records found for Month *${targetMonth}* of the current year.`);
    }
  }

  let textResponse = "";

  if (isTodayMode) {
    const r = results[results.length - 1]; // In case of duplicates, grab the latest row
    const loginReplacedText = r[4] === "TRUE" ? " _(Overwritten)_" : "";
    const logoutReplacedText = r[9] === "TRUE" ? " _(Overwritten)_" : "";
    const wfhText = r[7] === "TRUE" ? " 🏠 WFH" : "";
    
    textResponse = `📅 *Your Attendance for Today (${todayString})*\n\n` +
                   `*Login:* ${r[3] || "Missing"}${loginReplacedText}\n` +
                   `*Logout:* ${r[8] || "Not logged out yet"}${logoutReplacedText}\n` +
                   `*Hours Logged:* ${r[12] || "0"}\n` +
                   `*Status:* ${r[13] || "N/A"}${wfhText}`;
  } else {
    // Render a monospace ASCII table for monthly records
    textResponse = `📅 *Attendance Report for Month ${targetMonth}, ${currentYear}*\n`;
    textResponse += "```\n";
    textResponse += "Date  | In    | Out   | Hrs  | Status\n";
    textResponse += "---------------------------------------\n";
    
    results.forEach(r => {
      // Extract MM-DD and strip Slack emojis from the status for perfect table alignment
      const dateShort = String(r[1]).substring(5).padEnd(5); 
      const logIn = String(r[3] || "--:--").padEnd(5);
      const logOut = String(r[8] || "--:--").padEnd(5);
      const hrs = String(r[12] || "-").padEnd(4);
      const status = String(r[13] || "In Progress").replace(/[^\x20-\x7E]/g, "").trim().padEnd(11); 
      
      textResponse += `${dateShort} | ${logIn} | ${logOut} | ${hrs} | ${status}\n`;
    });
    
    textResponse += "```";
  }

  return sendEphemeralResponse(textResponse);
}

// =================================================================
// SPRINT 4: /LEAVE & /HELP COMMANDS
// =================================================================

function handleLeave(params, args) {
  const userId = params.user_id;
  const userName = getSlackRealName(userId) || params.user_name;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Attendance");

  const type = args.type ? String(args.type).toUpperCase() : null;
  const validTypes = {
    'A': 'Annual Leave', 'A1': 'First Half Annual Leave', 'A2': 'Second Half Annual Leave',
    'S': 'Sick Leave', 'S1': 'First Half Sick Leave', 'S2': 'Second Half Sick Leave'
  };

  if (!type || !validTypes[type]) {
    return sendEphemeralResponse("⚠️ *Invalid Leave Type!* Please specify `--type` using A, A1, A2, S, S1, or S2.");
  }

  const fromDateStr = resolveDate(args.from);
  const toDateStr = args.to ? resolveDate(args.to) : fromDateStr;

  const fromDate = new Date(fromDateStr);
  const toDate = new Date(toDateStr);

  if (fromDate > toDate) {
    return sendEphemeralResponse("⚠️ *Invalid Date Range!* The `--from` date cannot be after the `--to` date.");
  }

  const actualTimestamp = Utilities.formatDate(new Date(), "GMT+5", "yyyy-MM-dd HH:mm:ss");
  let appliedDays = 0;

  // Loop through the date range and generate a row for each leave day
  for (let d = new Date(fromDate); d <= toDate; d.setDate(d.getDate() + 1)) {
    const loopDateStr = Utilities.formatDate(d, "GMT+5", "yyyy-MM-dd");
    
    // Columns: A=Sr#, B=Date, C=Name, D=Login..M=TotalHrs, N=Status, O=LeaveCode, P=SlackID
    sheet.appendRow([
      "=ROW()-1", 
      loopDateStr, 
      userName, 
      "", "", actualTimestamp, "", false, 
      "", "", "", "", 0, 
      `On Leave ✈️ (${validTypes[type]})`, 
      type, 
      userId
    ]);
    appliedDays++;
  }

  return sendEphemeralResponse(`✅ *Leave Logged!* Successfully applied for *${validTypes[type]}* from *${fromDateStr}* to *${toDateStr}* (${appliedDays} day/s).`);
}

function handleHelp(params, args) {
  const rawText = (params.text || "").toLowerCase();

  // Detail view for specific commands
  if (rawText.includes("/login")) {
    return sendEphemeralResponse("🧑‍💻 *`/login` Command Details:*\n" +
      "> `--reason : ?` _(Optional, required if late)_\n" +
      "> `--replace : true | false` _(Optional, flags HR if overwritten)_\n" +
      "> `--date : YYYY-MM-DD | = | -n | +n` _(Optional, = means today)_\n" +
      "> `--time : HH:MM | =` _(Optional, = means current time)_\n" +
      "> `--wfh` _(Optional, if working from home)_");
  }
  
  if (rawText.includes("/logout")) {
    return sendEphemeralResponse("🏃 *`/logout` Command Details:*\n" +
      "> `--reason : ?` _(Optional, required if leaving early)_\n" +
      "> `--replace : true | false` _(Optional, flags HR if overwritten)_\n" +
      "> `--date : YYYY-MM-DD | = | -n | +n` _(Optional)_\n" +
      "> `--time : HH:MM | =` _(Optional)_");
  }

  if (rawText.includes("/attendance")) {
    return sendEphemeralResponse("📊 *`/attendance` Command Details:*\n" +
      "> `--[=]` _(Fetches today's ongoing attendance)_\n" +
      "> `--month 1-12 | = | -n | +n` _(Fetches full month report. Current year only.)_");
  }

  if (rawText.includes("/leave-req")) {
    return sendEphemeralResponse("✈️ *`/leave-req` Command Details:*\n" +
      "> `--type : A | A1 | A2 | S | S1 | S2` _(Required)_\n" +
      "> `--from : YYYY-MM-DD | = | -n | +n` _(Required)_\n" +
      "> `--to : YYYY-MM-DD | = | -n | +n` _(Optional, defaults to 'from' date)_");
  }

  // General list view (no explanations)
  return sendEphemeralResponse("🛠️ *Available Commands:*\n" +
    "`/login`\n" +
    "`/logout`\n" +
    "`/attendance`\n" +
    "`/leave-req`\n" +
    "`/help`\n\n" +
    "_Type `/help --/[command]` to see specific parameters (e.g., `/help --/leave-req`)._");
}