
const SHEETS = {
  master: "Task Master",
  employee: "Employee Dashboard",
  completion: "Completion Log",
  manager: "Manager Dashboard",
  alerts: "Alert & Escalation Log",
  settings: "Settings"
};

const TIME_ZONE = "Asia/Kolkata";

// ---------- SETUP ----------

function setupAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.values(SHEETS).forEach(name => {
    if (!ss.getSheetByName(name)) {
      throw new Error("Missing sheet: " + name);
    }
  });

  // Add a column to identify the exact task cycle in the alert log.
  const alertSheet = getSheet(SHEETS.alerts);
  if (!alertSheet.getRange(1, 11).getValue()) {
    alertSheet.getRange(1, 11).setValue("Task Due Date");
  }

  // Remove only this project's existing automation triggers.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (["onEditHandler", "runAutomation"].includes(
      trigger.getHandlerFunction()
    )) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onEditHandler")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger("runAutomation")
    .timeBased()
    .everyMinutes(5)
    .create();

  refreshDashboards();

  ss.toast("Automation setup completed!", "Setup", 5);
}

// ---------- COMMON FUNCTIONS ----------

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(name);

  if (!sheet) throw new Error("Missing sheet: " + name);

  return sheet;
}

function getSettings() {
  const values = getSheet(SHEETS.settings)
    .getRange("A2:B7")
    .getValues();

  const result = {};

  values.forEach(row => {
    if (row[0]) {
      result[String(row[0]).trim()] = row[1];
    }
  });

  return result;
}

function dateKey(value) {
  if (!value) return "";

  let d;

  if (value instanceof Date) {
    d = value;
  } else {
    const match = String(value).match(
      /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/
    );

    if (match) {
      d = new Date(
        Number(match[3]),
        Number(match[2]) - 1,
        Number(match[1])
      );
    } else {
      d = new Date(value);
    }
  }

  if (isNaN(d.getTime())) return "";

  return Utilities.formatDate(d, TIME_ZONE, "yyyy-MM-dd");
}

function addDays(dateValue, days) {
  const key = dateKey(dateValue);

  if (!key) throw new Error("Invalid task due date.");

  const parts = key.split("-");

  const d = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );

  d.setDate(d.getDate() + Number(days));

  return d;
}

function combineDateTime(dateValue, timeValue) {
  const key = dateKey(dateValue);

  if (!key) return null;

  const parts = key.split("-");

  const d = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );

  if (timeValue instanceof Date) {
    d.setHours(
      timeValue.getHours(),
      timeValue.getMinutes(),
      0,
      0
    );
  } else if (typeof timeValue === "number") {
    const totalMinutes = Math.round(timeValue * 1440);

    d.setHours(
      Math.floor(totalMinutes / 60),
      totalMinutes % 60,
      0,
      0
    );
  } else if (
    typeof timeValue === "string" &&
    timeValue.includes(":")
  ) {
    const parts = timeValue.split(":");

    d.setHours(
      Number(parts[0]),
      Number(parts[1]),
      0,
      0
    );
  } else {
    return null;
  }

  return d;
}

// ---------- DASHBOARDS ----------

function refreshDashboards() {
  refreshEmployeeDashboard();
  refreshManagerDashboard();
}

function refreshEmployeeDashboard() {
  const master = getSheet(SHEETS.master);
  const dashboard = getSheet(SHEETS.employee);

  const lastDashboardRow = dashboard.getLastRow();

  const oldRows = lastDashboardRow > 1
    ? dashboard.getRange(
        2, 1, lastDashboardRow - 1, 9
      ).getValues()
    : [];

  const oldMap = {};

  oldRows.forEach(row => {
    const key = String(row[0]) + "|" + dateKey(row[3]);
    oldMap[key] = row;
  });

  const output = [];
  const lastMasterRow = master.getLastRow();

  if (lastMasterRow >= 2) {
    const tasks = master.getRange(
      2, 1, lastMasterRow - 1, 12
    ).getValues();

    tasks.forEach(task => {
      if (
        !task[0] ||
        String(task[11]).toUpperCase() !== "TRUE"
      ) return;

      const key = String(task[0]) + "|" + dateKey(task[7]);
      const old = oldMap[key];

      output.push([
        task[0],
        task[1],
        task[2],
        task[7],
        task[6],
        old ? old[5] || "Pending" : "Pending",
        old ? old[6] || "" : "",
        old ? old[7] || "" : "",
        old ? old[8] || "" : ""
      ]);
    });
  }

  dashboard.getRange("A2:I1000").clearContent();

  if (output.length) {
    dashboard.getRange(
      2, 1, output.length, 9
    ).setValues(output);
  }
}

function refreshManagerDashboard() {
  const master = getSheet(SHEETS.master);
  const dashboard = getSheet(SHEETS.manager);

  const lastDashboardRow = dashboard.getLastRow();

  const oldRows = lastDashboardRow > 1
    ? dashboard.getRange(
        2, 1, lastDashboardRow - 1, 9
      ).getValues()
    : [];

  const oldMap = {};

  oldRows.forEach(row => {
    const key = String(row[0]) + "|" + dateKey(row[3]);
    oldMap[key] = row;
  });

  const output = [];
  const lastMasterRow = master.getLastRow();

  if (lastMasterRow >= 2) {
    const tasks = master.getRange(
      2, 1, lastMasterRow - 1, 12
    ).getValues();

    tasks.forEach(task => {
      if (
        !task[0] ||
        String(task[11]).toUpperCase() !== "TRUE"
      ) return;

      const key = String(task[0]) + "|" + dateKey(task[7]);
      const old = oldMap[key];

      const due = combineDateTime(task[7], task[6]);

      const delay = due
        ? Math.max(
            0,
            Math.floor((new Date() - due) / 60000)
          )
        : 0;

      output.push([
        task[0],
        task[1],
        task[2],
        task[7],
        task[6],
        getTaskStatus(task[0], task[7]),
        delay,
        old ? old[7] || "Pending" : "Pending",
        old ? old[8] || "Not Escalated" : "Not Escalated"
      ]);
    });
  }

  dashboard.getRange("A2:I1000").clearContent();

  if (output.length) {
    dashboard.getRange(
      2, 1, output.length, 9
    ).setValues(output);
  }
}

function getTaskStatus(taskId, dueDate) {
  const rows = getSheet(SHEETS.employee)
    .getDataRange()
    .getValues();

  for (let i = 1; i < rows.length; i++) {
    if (
      String(rows[i][0]) === String(taskId) &&
      dateKey(rows[i][3]) === dateKey(dueDate)
    ) {
      return rows[i][5] || "Pending";
    }
  }

  return "Pending";
}

// ---------- EDIT HANDLER ----------

function onEditHandler(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row < 2) return;

  if (
    name === SHEETS.employee &&
    col === 6 &&
    e.range.getNumRows() === 1
  ) {
    const status = String(
      sheet.getRange(row, 6).getValue()
    ).toLowerCase();

    if (status === "completed") {
      recordCompletion(row);
    }
  }

  if (name === SHEETS.manager && col === 8) {
    updateAcknowledgement(row);
  }

  if (name === SHEETS.manager && col === 9) {
    updateEscalation(row);
  }
}

// ---------- COMPLETION ----------

function recordCompletion(row) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) return;

  try {
    const dashboard = getSheet(SHEETS.employee);
    const master = getSheet(SHEETS.master);
    const log = getSheet(SHEETS.completion);

    const values = dashboard.getRange(
      row, 1, 1, 9
    ).getValues()[0];

    const taskId = values[0];
    const taskName = values[1];
    const assignedTo = values[2];
    const dueDate = values[3];
    const proof = values[6];

    if (!taskId || !dueDate) return;

    if (!proof) {
      dashboard.getRange(row, 6).setValue("Pending");

      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Enter completion proof in column G first.",
        "Proof required",
        5
      );

      return;
    }

    const tasks = master.getDataRange().getValues();

    let masterRow = -1;
    let frequencyDays = 1;
    let currentDueDate = null;

    for (let i = 1; i < tasks.length; i++) {
      if (String(tasks[i][0]) === String(taskId)) {
        masterRow = i + 1;
        frequencyDays = Number(tasks[i][4]) || 1;
        currentDueDate = tasks[i][7];
        break;
      }
    }

    if (masterRow === -1) {
      throw new Error("Task not found in Task Master.");
    }

    // Prevent completion of an outdated task cycle.
    if (dateKey(currentDueDate) !== dateKey(dueDate)) {
      refreshDashboards();
      return;
    }

    // Prevent duplicate completion for the same task cycle.
    const existing = log.getDataRange().getValues().some(entry =>
      String(entry[1]) === String(taskId) &&
      dateKey(entry[4]) === dateKey(dueDate)
    );

    if (existing) {
      refreshDashboards();
      return;
    }

    const completedBy = values[7] || assignedTo;
    const completedAt = new Date();

    // Save the completed cycle.
    log.appendRow([
      "LOG-" + Utilities.getUuid().slice(0, 8),
      taskId,
      taskName,
      assignedTo,
      dueDate,
      proof,
      completedBy,
      completedAt,
      dueDate
    ]);

    // Calculate the next recurring cycle from the due date.
    const nextDue = addDays(dueDate, frequencyDays);

    master.getRange(masterRow, 8).setValue(nextDue);
    master.getRange(masterRow, 9).setValue(completedAt);
    master.getRange(masterRow, 10).setValue(completedBy);

    // The dashboard now displays the new cycle as Pending.
    refreshDashboards();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Completion saved. Next cycle created.",
      "Task completed",
      5
    );

  } finally {
    lock.releaseLock();
  }
}

// ---------- AUTOMATIC CHECK ----------

function runAutomation() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) return;

  try {
    // Do not refresh the Employee Dashboard on a timer.
    // This protects employee proof entries.
    refreshManagerDashboard();
    processAlerts();

  } finally {
    lock.releaseLock();
  }
}

// ---------- ALERTS AND ESCALATION ----------

function processAlerts() {
  const settings = getSettings();

  const managerEmail = String(
    settings["Manager Email"] || ""
  ).trim();

  const leadershipEmail = String(
    settings["Leadership Email"] || ""
  ).trim();

  const alertDelay = Number(
    settings["Manager Alert Delay (Minutes)"]
  ) || 30;

  const ackSla = Number(
    settings["Manager Acknowledgement SLA (Hours)"]
  ) || 2;

  const escalationDelay = Number(
    settings["Leadership Escalation Delay (Hours)"]
  ) || 4;

  const master = getSheet(SHEETS.master);
  const alerts = getSheet(SHEETS.alerts);
  const manager = getSheet(SHEETS.manager);
  const employee = getSheet(SHEETS.employee);

  const tasks = master.getDataRange().getValues();
  const now = new Date();

  tasks.slice(1).forEach(task => {
    const taskId = task[0];
    const taskName = task[1];
    const assignedTo = task[2];
    const dueTime = task[6];
    const dueDate = task[7];
    const active = task[11];

    if (
      !taskId ||
      String(active).toUpperCase() !== "TRUE"
    ) return;

    const employeeRows = employee.getDataRange().getValues();

    const completed = employeeRows.some(row =>
      String(row[0]) === String(taskId) &&
      dateKey(row[3]) === dateKey(dueDate) &&
      String(row[5]).toLowerCase() === "completed"
    );

    if (completed) return;

    const due = combineDateTime(dueDate, dueTime);

    if (!due) return;

    const overdueMinutes = (now - due) / 60000;

    if (overdueMinutes < alertDelay) return;

    let alertRows = alerts.getDataRange().getValues();

    const managerAlert = alertRows.find(row =>
      String(row[1]) === String(taskId) &&
      row[4] === "Manager Alert" &&
      dateKey(row[10]) === dateKey(dueDate)
    );

    if (!managerAlert && managerEmail) {
      sendAlert(
        alerts,
        taskId,
        taskName,
        assignedTo,
        "Manager Alert",
        managerEmail,
        "Task overdue: " + taskName,
        "Task " + taskId +
        " assigned to " + assignedTo +
        " is overdue. Please review the Manager Dashboard.",
        dueDate
      );

      alertRows = alerts.getDataRange().getValues();
    }

    const currentManagerAlert = alertRows.find(row =>
      String(row[1]) === String(taskId) &&
      row[4] === "Manager Alert" &&
      dateKey(row[10]) === dateKey(dueDate)
    );

    if (!currentManagerAlert) return;

    const alertTime = new Date(currentManagerAlert[5]);

    const managerRows = manager.getDataRange().getValues();

    const matchingManagerRow = managerRows.find(row =>
      String(row[0]) === String(taskId) &&
      dateKey(row[3]) === dateKey(dueDate)
    );

    const acknowledged = matchingManagerRow &&
      String(matchingManagerRow[7]).toLowerCase() === "acknowledged";

    const hoursSinceAlert = (
      now - alertTime
    ) / 3600000;

    const shouldEscalate =
      (!acknowledged && hoursSinceAlert >= ackSla) ||
      hoursSinceAlert >= escalationDelay;

    const alreadyEscalated = alertRows.some(row =>
      String(row[1]) === String(taskId) &&
      row[4] === "Leadership Escalation" &&
      dateKey(row[10]) === dateKey(dueDate)
    );

    if (
      shouldEscalate &&
      !alreadyEscalated &&
      leadershipEmail
    ) {
      sendAlert(
        alerts,
        taskId,
        taskName,
        assignedTo,
        "Leadership Escalation",
        leadershipEmail,
        "Escalation: unresolved task " + taskName,
        "Task " + taskId +
        " remains unresolved. Please review.",
        dueDate
      );

      updateManagerEscalation(taskId, dueDate);
    }
  });
}

function sendAlert(
  sheet,
  taskId,
  taskName,
  assignedTo,
  type,
  recipient,
  subject,
  body,
  dueDate
) {
  const alertId = "ALT-" + Utilities.getUuid().slice(0, 8);
  const sentAt = new Date();

  MailApp.sendEmail(recipient, subject, body);

  sheet.appendRow([
    alertId,
    taskId,
    taskName,
    assignedTo,
    type,
    sentAt,
    recipient,
    "Pending",
    type === "Leadership Escalation"
      ? "Escalated"
      : "Not Escalated",
    type === "Leadership Escalation"
      ? sentAt
      : "",
    dueDate
  ]);
}

// ---------- MANAGER ACKNOWLEDGEMENT ----------

function updateAcknowledgement(row) {
  const manager = getSheet(SHEETS.manager);
  const alerts = getSheet(SHEETS.alerts);

  const task = manager.getRange(
    row, 1, 1, 9
  ).getValues()[0];

  if (
    String(task[7]).toLowerCase() !== "acknowledged"
  ) return;

  const rows = alerts.getDataRange().getValues();

  rows.forEach((alert, i) => {
    if (
      String(alert[1]) === String(task[0]) &&
      alert[4] === "Manager Alert" &&
      dateKey(alert[10]) === dateKey(task[3])
    ) {
      alerts.getRange(i + 1, 8).setValue("Acknowledged");
    }
  });
}

// ---------- ESCALATION STATUS ----------

function updateEscalation(row) {
  const manager = getSheet(SHEETS.manager);

  const task = manager.getRange(
    row, 1, 1, 9
  ).getValues()[0];

  if (
    String(task[8]).toLowerCase() === "escalated"
  ) {
    updateManagerEscalation(task[0], task[3]);
  }
}

function updateManagerEscalation(taskId, dueDate) {
  const manager = getSheet(SHEETS.manager);
  const alerts = getSheet(SHEETS.alerts);

  const managerRows = manager.getDataRange().getValues();

  managerRows.forEach((row, i) => {
    if (
      String(row[0]) === String(taskId) &&
      dateKey(row[3]) === dateKey(dueDate)
    ) {
      manager.getRange(i + 1, 9).setValue("Escalated");
    }
  });

  const alertRows = alerts.getDataRange().getValues();

  alertRows.forEach((row, i) => {
    if (
      String(row[1]) === String(taskId) &&
      row[4] === "Manager Alert" &&
      dateKey(row[10]) === dateKey(dueDate)
    ) {
      alerts.getRange(i + 1, 9).setValue("Escalated");
    }
  });
}