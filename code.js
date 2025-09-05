// ✅ 日期欄位一定要判斷是否為 Date 並轉為 "yyyy/MM/dd" 格式再進行比較
// 否則會導致 == 比對失敗、條件永遠不成立 顯示用途也要處理日期格式，避免出現 GMT/UTC 雜訊。
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}
function getProfileData() {
  const ss = SpreadsheetApp.openById('16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s');
  const sheet = ss.getSheetByName('Profile');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  const profile = {};
  headers.forEach((key, i) => profile[key] = row[i]);

  const birthdayFormatted = (profile.birthday instanceof Date)
    ? Utilities.formatDate(profile.birthday, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : (profile.birthday || '');

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const lastSurveyDateFormatted = (profile.LastSurveyDate instanceof Date)
    ? Utilities.formatDate(profile.LastSurveyDate, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : (profile.LastSurveyDate || '');
  const surveyFilledToday = (lastSurveyDateFormatted === today);

  return {
    playerName: profile.PlayerName,
    birthday: birthdayFormatted,
    coins: profile.Coins,
    honorPoints: profile.HonorPoints,
    todayActionPoints: profile.TodayActionPoints,
    cleanliness: profile.Cleanliness,
    mood: profile.Mood,
    energy: profile.Energy,
    health: profile.Health,
    selfDiscipline: profile.SelfDiscipline, 
    backgroundUrl: "https://renewmiffy.github.io/mygame/img/bg/" + profile.currentBackground,
    characterUrl: "https://renewmiffy.github.io/mygame/img/char/" + profile.currentCharacter,
    Intelligence: profile.Intelligence,
    IntelligenceLevel: profile.IntelligenceLevel,
    Physical: profile.Physical,
    PhysicalLevel: profile.PhysicalLevel,
    Sensitivity: profile.Sensitivity,
    SensitivityLevel: profile.SensitivityLevel,
    Creativity: profile.Creativity,
    CreativityLevel: profile.CreativityLevel,
    Grace: profile.Grace,
    GraceLevel: profile.GraceLevel,
    surveyFilledToday: surveyFilledToday
  };
}
function getSurveyQuestions() {
  const ss = SpreadsheetApp.openById('16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s');
  const sheet = ss.getSheetByName('SurveyQuestions');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows
    .filter(row => row[headers.indexOf('Active')] === true)
    .map(row => {
      const result = {};
      headers.forEach((key, i) => {
        result[key] = row[i];
      });
      return result;
    });
}


function handleSurvey(formData) {
  const ss = SpreadsheetApp.openById('16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s');
  const profileSheet = ss.getSheetByName('Profile');
  const questionsSheet = ss.getSheetByName('SurveyQuestions');
  const logSheet = ss.getSheetByName('SurveyLog');
  const rulesSheet = ss.getSheetByName('ScreenTimeRules');

  const headers = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const row = profileSheet.getRange(2, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profile = {};
  headers.forEach((key, i) => profile[key] = row[i]);

  const today = new Date();
  const surveyAnswers = [];
  const questions = questionsSheet.getRange(2, 1, questionsSheet.getLastRow() - 1, 6).getValues()
    .filter(r => r[5] === true);

  questions.forEach(q => {
    const qkey = q[0];
    const value = formData[qkey];
    surveyAnswers.push(value);

    try {
      const optionsMap = JSON.parse(q[4]);
      const effectObj = optionsMap[value];
      if (!effectObj) return;

      Object.entries(effectObj).forEach(([field, effectValue]) => {
        const current = parseFloat(profile[field] || 0);
        profile[field] = (typeof effectValue === 'string' && /^[+-]/.test(effectValue))
          ? current + parseFloat(effectValue)
          : parseFloat(effectValue);
      });
    } catch (e) {}
  });

  try {
    const rules = rulesSheet.getDataRange().getValues();
    const screenTime = Math.max(0, parseFloat(formData["Q1_ScreenTime"] || 0));
    const isHoliday = formData["Q0_IsHoliday"] == 1 ? 1 : 0;

    for (let i = 1; i < rules.length; i++) {
      const [ruleHoliday, min, max, field, value] = rules[i];
      if (ruleHoliday === isHoliday && screenTime >= min && screenTime <= max) {
        profile[field] = parseFloat(profile[field] || 0) + parseFloat(value);
      }
    }
  } catch (e) {}

  const baseAP = 2;
  const energyAP = Math.min(parseFloat(profile.Energy || 0), 100) / 20;  // 最高給 5 點 A
  const selfDisciplineAP = Math.min(parseFloat(profile.SelfDiscipline || 0), 1) * 2;
  const dailyDoneAP = parseFloat(profile.DailyTaskDoneCount || 0) * 0.2;
  profile.TodayActionPoints = Math.floor(baseAP + energyAP + selfDisciplineAP + dailyDoneAP);
  profile.Cleanliness = (parseFloat(profile.Cleanliness) || 0) - 20;
  profile.LastSurveyDate = today;

  writeProfile(profile, "問卷");

  const logRow = [today, ...surveyAnswers, new Date()];
  logSheet.appendRow(logRow);
}

function getRecentLogs() {
  const ss = SpreadsheetApp.openById('16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s');
  const logSheet = ss.getSheetByName('PlayerLog');
  const mapSheet = ss.getSheetByName('FieldMapping');

  const lastRow = logSheet.getLastRow();
  const lastCol = logSheet.getLastColumn();
  if (lastRow <= 1) return [];

  const mapRows = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues();
  const fieldMap = {};
  mapRows.forEach(r => { fieldMap[r[1]] = r[0]; });

  const headers = logSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows = logSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return rows.slice(-10).map(row => {
    const log = {};
    headers.forEach((key, i) => {
      if (key === 'CreatedAt' && row[i] instanceof Date) {
        log[key] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), 'MM/dd HH:mm');
      } else {
    log[key] = row[i];
}

    });
    return log;
  });
}

function applyAttributeLimit(profile) {
  const limits = {
    Cleanliness: [0, 100],
    Mood: [0, 100],
    Energy: [0, 100],
    Health: [0, 100],
    SelfDiscipline: [0, 100],
    Intelligence: [0, 999],
    Physical: [0, 999],
    Creativity: [0, 999],
    Sensitivity: [0, 999],
    Grace: [0, 999],
    TodayActionPoints: [0, 99]
  };

  Object.entries(limits).forEach(([field, [min, max]]) => {
    const val = parseFloat(profile[field] ?? 0);
    profile[field] = Math.min(max, Math.max(min, isNaN(val) ? 0 : val));
  });
}

function writeProfile(profile, source = "系統") {
  const ss = SpreadsheetApp.openById("16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s");
  const profileSheet = ss.getSheetByName("Profile");
  const logSheet = ss.getSheetByName("PlayerLog");

  const headers = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const oldRow = profileSheet.getRange(2, 1, 1, headers.length).getValues()[0];
  const oldProfile = {};
  headers.forEach((key, i) => oldProfile[key] = oldRow[i]);

  applyAttributeLimit(profile);

  const newRow = headers.map(k => profile[k] ?? '');
  profileSheet.getRange(2, 1, 1, headers.length).setValues([newRow]);

  // ✅ 關鍵屬性欄位（才寫入 log）
  const fieldsToLog = [
    "Cleanliness", "Mood", "Energy", "Health", "SelfDiscipline",
    "Intelligence", "Physical", "Creativity", "Sensitivity", "Grace",
    "Coins", "HonorPoints", "TodayActionPoints"
  ];

  const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const now = new Date();
  const logsToWrite = [];

  fieldsToLog.forEach(key => {
    const oldVal = parseFloat(oldProfile[key] ?? 0);
    const newVal = parseFloat(profile[key] ?? 0);
    if (!isNaN(oldVal) && !isNaN(newVal) && oldVal !== newVal) {
      const delta = newVal - oldVal;
      const formattedDelta = Math.round(delta * 100) / 100;  // 四捨五入到小數第 2 位

      const logRow = {
        CreatedAt: now,
        Type: "action",
        ActionName: source,
        AffectField: key,
        AffectValue: formattedDelta,
        Source: source
};

      logsToWrite.push(logHeaders.map(k => logRow[k] ?? ''));
    }
  });

  if (logsToWrite.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToWrite.length, logHeaders.length)
      .setValues(logsToWrite);
  }
}






// ✅ 這是後端 Apps Script 用的
function getCurrentStatus() {
  return [];
}
// ✅ getMissionList()：支援獎勵 JSON 格式與 LastClaimedDate 判斷
function getMissionList() {
  const ss = SpreadsheetApp.getActive();
  const missionSheet = ss.getSheetByName("MissionCenter");
  const profileSheet = ss.getSheetByName("Profile");
  const doneTasksSheet = ss.getSheetByName("DailyTasks");

  const missionData = missionSheet.getDataRange().getValues();
  const missionHeaders = missionData[0];
  const missionRows = missionData.slice(1);

  const getMissionField = (row, field) => {
    const idx = missionHeaders.indexOf(field);
    if (idx === -1) throw new Error(`❌ 找不到欄位：「${field}」`);
    return row[idx];
  };

  const profileRow = profileSheet.getDataRange().getValues()[1];
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileRow.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((key, i) => profile[key] = profileRow[i]);

  const today = formatYMD(new Date());
  const dailyData = doneTasksSheet.getDataRange().getValues().slice(1);

  const todayTaskIds = dailyData
    .filter(row => formatYMD(row[3]) === today)
    .map(row => row[0]);

  const totalTaskDone = dailyData.filter(row => row[5] > 0).length;

  const result = missionRows.map((row, i) => {
    const missionId = getMissionField(row, "MissionID");
    const name = getMissionField(row, "任務名稱");
    const conditionType = getMissionField(row, "條件類型");
    const param = getMissionField(row, "條件參數");
    const type = getMissionField(row, "類型") || "Daily";
    const repeatable = getMissionField(row, "可重複") === "是";
    const displayOrder = parseInt(getMissionField(row, "顯示順序") || 0);
    const lastClaimed = getMissionField(row, "LastClaimedDate");

    let fulfilled = false;
    if (conditionType === "DailyTaskDoneCount") {
      fulfilled = totalTaskDone >= parseInt(param);
    } else if (conditionType === "TaskDoneToday") {
      fulfilled = todayTaskIds.includes(param);
    } else if (conditionType === "TaskDoneCount") {
      const [taskId, count] = param.split(":");
      const taskRow = dailyData.find(r => r[0] === taskId);
      fulfilled = taskRow && parseInt(taskRow[5] || 0) >= parseInt(count);
    }

    let claimed = false;
    if (type === "Daily") {
      claimed = formatYMD(lastClaimed) === today;
    } else {
      claimed = !!lastClaimed && !repeatable;
    }

    let rewardText = "";
    try {
      const rewardObj = JSON.parse(getMissionField(row, "獎勵") || "{}");
      rewardText = Object.entries(rewardObj)
        .map(([key, val]) => `+${val} ${key}`)
        .join(" / ");
    } catch (e) {
      rewardText = "❌ 獎勵格式錯誤";
    }

    return {
      id: missionId,
      name: name,
      description: `🎯 ${name}`,
      rewardText: rewardText,
      fulfilled: fulfilled,
      claimed: claimed,
      displayOrder: displayOrder,
      rowIndex: i,
      type: type // 加入類型給一鍵領取判斷用
    };
  });

  return result.sort((a, b) => a.displayOrder - b.displayOrder);
}


// ✅ claimDailyTask()：支援 JSON 獎勵 + LastClaimedDate 寫入
function claimDailyTask(missionId) {
  const ss = SpreadsheetApp.getActive();
  const missionSheet = ss.getSheetByName("MissionCenter");
  const profileSheet = ss.getSheetByName("Profile");

  const missionData = missionSheet.getDataRange().getValues();
  const missionHeaders = missionData[0];
  const missionRows = missionData.slice(1);

  const getMissionField = (row, field) => {
    const idx = missionHeaders.indexOf(field);
    return row[idx];
  };

  const missionIndex = missionRows.findIndex(row => getMissionField(row, "MissionID") === missionId);
  if (missionIndex === -1) throw new Error("❌ 找不到指定任務");
  const mission = missionRows[missionIndex];

  const profileRange = profileSheet.getRange(2, 1, 1, profileSheet.getLastColumn());
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileRange.getValues()[0];
  const profile = {};
  profileHeaders.forEach((key, i) => profile[key] = profileRow[i]);

  const repeatable = getMissionField(mission, "可重複") === "是";
  const lastClaimed = getMissionField(mission, "LastClaimedDate");
  const type = getMissionField(mission, "類型") || "Daily";
  const today = formatYMD(new Date());

  if (type === "Daily" && formatYMD(lastClaimed) === today) {
    throw new Error("⚠️ 今日已領取");
  }
  if (type !== "Daily" && lastClaimed && !repeatable) {
    throw new Error("⚠️ 此任務已領取");
  }

  const rewardJson = getMissionField(mission, "獎勵");
  let rewardText = "";

  try {
    const rewardObj = JSON.parse(rewardJson);
    for (const [field, value] of Object.entries(rewardObj)) {
      if (!profile.hasOwnProperty(field)) {
        throw new Error(`⚠️ Profile 表未定義欄位：${field}`);
      }
      profile[field] = (parseFloat(profile[field]) || 0) + parseFloat(value);
      rewardText += `+${value} ${field} `;
    }
  } catch (e) {
    throw new Error("❌ 無法解析任務獎勵：" + e.message);
  }

  const updatedRow = profileHeaders.map(k => profile[k] ?? '');
  profileRange.setValues([updatedRow]);

  // ✅ 寫入任務表的 LastClaimedDate
  const colIndex = missionHeaders.indexOf("LastClaimedDate") + 1;
  if (colIndex > 0) {
    missionSheet.getRange(missionIndex + 2, colIndex).setValue(today);
  }

  return `✅ 已領取 ${getMissionField(mission, "任務名稱")}：${rewardText.trim()}`;
}



// ✅ 工具函式：統一處理所有日期欄位 ➜ yyyy/MM/dd 字串
function formatYMD(value) {
  return (value instanceof Date)
    ? Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : value;
}

function doDailyTask(taskId) {
  const ss = SpreadsheetApp.openById("16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s");
  const taskSheet = ss.getSheetByName("DailyTasks");
  const profileSheet = ss.getSheetByName("Profile");
  const logSheet = ss.getSheetByName("PlayerLog");

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

  const taskData = taskSheet.getDataRange().getValues();
  const taskHeaders = taskData[0];
  const taskRows = taskData.slice(1);
  const taskIndex = taskRows.findIndex(r => r[0] === taskId);
  if (taskIndex === -1) throw new Error("❌ 找不到此任務");

  const taskRow = taskRows[taskIndex];
  const getTaskField = (field) => taskRow[taskHeaders.indexOf(field)];
  const lastDate = getTaskField("LastDoneDate");

  if (lastDate instanceof Date && Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "yyyy/MM/dd") === todayStr) {
    throw new Error("⚠️ 此任務今日已完成");
  }

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((key, i) => profile[key] = profileRow[i]);

  const effects = JSON.parse(getTaskField("Effects"));
  Object.entries(effects).forEach(([field, value]) => {
    const current = parseFloat(profile[field] || 0);
    profile[field] = current + parseFloat(value);
  });

  // ✅ 新增具體任務名稱記錄
  const taskName = getTaskField("任務名稱") || "未知任務";
  writeProfile(profile, "日常任務 - " + taskName);

  const today = new Date();
  const taskRowIndex = taskIndex + 2;
  const lastDoneDateCell = taskSheet.getRange(taskRowIndex, taskHeaders.indexOf("LastDoneDate") + 1);
  const streakCell = taskSheet.getRange(taskRowIndex, taskHeaders.indexOf("StreakCount") + 1);
  const totalCell = taskSheet.getRange(taskRowIndex, taskHeaders.indexOf("TotalDoneCount") + 1);

  const yesterdayStr = Utilities.formatDate(new Date(today.getTime() - 86400000), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const lastDateStr = (lastDate instanceof Date)
    ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : "";

  const streakCount = parseInt(getTaskField("StreakCount") || 0);
  const totalCount = parseInt(getTaskField("TotalDoneCount") || 0);

  const newStreak = (lastDateStr === yesterdayStr) ? streakCount + 1 : 1;
  lastDoneDateCell.setValue(today);
  streakCell.setValue(newStreak);
  totalCell.setValue(totalCount + 1);
}


function getDailyTaskList() {
  const ss = SpreadsheetApp.openById("16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s");
  const sheet = ss.getSheetByName("DailyTasks");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);

    let lastDateStr = "";
    if (obj.LastDoneDate instanceof Date) {
      lastDateStr = Utilities.formatDate(obj.LastDoneDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
    }

    return {
      id: obj.TaskID,
      name: obj.任務名稱,
      effects: obj.Effects,
      fulfilledToday: (lastDateStr === todayStr)
    };
  });
}
function doLearning(skillId) {
  const ss = SpreadsheetApp.openById("16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s");
  const skillSheet = ss.getSheetByName("SkillMaster");
  const profileSheet = ss.getSheetByName("Profile");

  if (!skillSheet) throw new Error("❌ 找不到 SkillMaster 表");

  const data = skillSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const idx = rows.findIndex(r => r[0] === skillId);
  if (idx === -1) throw new Error("❌ 找不到此技能");

  const row = rows[idx];
  const get = (field) => row[headers.indexOf(field)];
  const effects = JSON.parse(get("Effects") || "{}");

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const lastDate = get("LastDoneDate");
  const lastDateStr = (lastDate instanceof Date)
    ? Utilities.formatDate(lastDate, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : "";

  if (lastDateStr === todayStr) throw new Error("⚠️ 今天已學過此技能");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  Object.entries(effects).forEach(([field, value]) => {
    profile[field] = (parseFloat(profile[field]) || 0) + parseFloat(value);
  });

  // ✅ 加入具體技能名稱
  const skillName = get("SkillName") || "未知技能";
  writeProfile(profile, "學習 - " + skillName);

  const sheetRow = idx + 2;
  const today = new Date();
  const streakIndex = headers.indexOf("StreakCount");
  const totalIndex = headers.indexOf("TotalDoneCount");
  const dateIndex = headers.indexOf("LastDoneDate");

  const yesterdayStr = Utilities.formatDate(new Date(today.getTime() - 86400000), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const newStreak = (lastDateStr === yesterdayStr) ? (parseInt(get("StreakCount")) || 0) + 1 : 1;
  const newTotal = (parseInt(get("TotalDoneCount")) || 0) + 1;

  skillSheet.getRange(sheetRow, dateIndex + 1).setValue(today);
  skillSheet.getRange(sheetRow, streakIndex + 1).setValue(newStreak);
  skillSheet.getRange(sheetRow, totalIndex + 1).setValue(newTotal);
}



function getAllSkills() {
  const ss = SpreadsheetApp.openById("16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s");
  const sheet = ss.getSheetByName("SkillMaster");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(row => {
    const skill = {};
    headers.forEach((h, i) => {
      if (h === "LastDoneDate" && row[i] instanceof Date) {
        skill[h] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else {
        skill[h] = row[i];
      }
    });
    return skill;
  });
}
function getAttributeMap() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("FieldMapping");
  const values = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const [chinese, english] = values[i];
    if (chinese && english && english.endsWith("Boost")) { // 限定只抓屬性加成欄
      map[english] = chinese;
    }
  }

  return map;
}
function getEquippedInfo() {
  const ss = SpreadsheetApp.getActive();
  const profileSheet = ss.getSheetByName("Profile");
  const equipSheet = ss.getSheetByName("EquipmentMaster");
  const mapSheet = ss.getSheetByName("FieldMapping");
  const attrSheet = ss.getSheetByName("AdventureAttributes");

  // 取得 Profile 資料
  const profileRow = profileSheet.getRange(2, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileHeader = profileSheet.getRange(1, 1, 1, profileRow.length).getValues()[0];
  const profile = {};
  profileHeader.forEach((k, i) => profile[k] = profileRow[i]);

  const playerName = profile["PlayerName"];
  if (!playerName) throw new Error("❌ Profile 中找不到 PlayerName");

  // 抓裝備
  const equippedSlots = ['Head', 'Body', 'Gloves', 'Belt', 'Shoes', 'Accessory','Weapon'];
  const equipped = {}, equipIDs = [];
  equippedSlots.forEach(slot => {
    const id = profile[slot];
    equipped[slot] = id || '';
    if (id) equipIDs.push(id);
  });

  // 抓裝備加成資料
  const equipData = equipSheet.getDataRange().getValues();
  const equipHeader = equipData[0];
  const equipRows = equipData.slice(1).filter(row => equipIDs.includes(row[0]));

  // 抓屬性欄位與值（AdventureAttributes）
  const attrData = attrSheet.getDataRange().getValues();
  const attrHeader = attrData[0];
  const attrRow = attrData.find(r => r[0] === playerName);
  if (!attrRow) throw new Error("❌ AdventureAttributes 中找不到玩家資料");

  // 抓欄位名稱對照（FieldMapping）
  const mapRows = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues();
  const attrNames = {};
  mapRows.forEach(([zh, en]) => {
    if (zh && en) {
      attrNames[en] = zh;
    }
  });

  // 對應 slot 名稱
  const slotNames = {};
  equippedSlots.forEach(slot => {
    const zh = mapRows.find(r => r[1] === slot)?.[0] || slot;
    slotNames[slot] = zh;
  });

  // 對應裝備顯示名稱
  const nameMap = {};
  equipRows.forEach(row => {
    nameMap[row[0]] = row[equipHeader.indexOf("Name")];
  });
  Object.entries(equipped).forEach(([slot, id]) => {
    equipped[slot] = nameMap[id] || '空';
  });

  // 🚀 抓取 AdventureAttributes 中所有 BaseXXX 欄位
  const attributes = {};
  attrHeader.forEach((field, i) => {
    if (!/^Base[A-Z]/.test(field)) return; // 只處理 Base 開頭欄位

    const attrKey = field.replace(/^Base/, '');  // 例如 BaseHP ➜ HP
    const base = parseFloat(attrRow[i]) || 0;

    const equipKey = "Equip" + attrKey;
    const bonusKey = attrKey + "Bonus";
    const finalKey = "Final" + attrKey;

    const equipIndex = attrHeader.indexOf(equipKey);
    const bonusIndex = profileHeader.indexOf(bonusKey);
    const finalIndex = attrHeader.indexOf(finalKey);

    const equip = (equipIndex !== -1) ? parseFloat(attrRow[equipIndex]) || 0 : 0;
    const bonus = (bonusIndex !== -1) ? parseFloat(profile[bonusKey]) || 0 : 0;
    const final = (finalIndex !== -1) ? parseFloat(attrRow[finalIndex]) || (base + equip + bonus) : (base + equip + bonus);

    attributes[attrKey] = {
      base,
      equip,
      bonus,
      final
    };
  });

  // ✅ 回傳給前端使用
  return {
    playerName,
    equipped,
    slotNames,
    attrNames,
    attributes
  };
}

function equipItem(slot, equipmentID) {
  const ss = SpreadsheetApp.getActive();
  const profileSheet = ss.getSheetByName("Profile");
  const inventorySheet = ss.getSheetByName("Inventory");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];

  // 先找出欄位位置
  const slotCol = profileHeaders.indexOf(slot);
  if (slotCol === -1) throw new Error("❌ Profile 中找不到欄位: " + slot);

  // 把舊裝備設為未裝備
  const oldEquipID = profileRow[slotCol];
  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const idCol = invHeaders.indexOf("ItemID");
  const equippedCol = invHeaders.indexOf("Equipped");

  for (let i = 1; i < invData.length; i++) {
    const row = invData[i];
    if (row[idCol] === oldEquipID) {
      inventorySheet.getRange(i + 1, equippedCol + 1).setValue(false);
    }
  }

  // 把新裝備設為已裝備
  for (let i = 1; i < invData.length; i++) {
    const row = invData[i];
    if (row[idCol] === equipmentID) {
      inventorySheet.getRange(i + 1, equippedCol + 1).setValue(true);
      break;
    }
  }

  // 更新 profile 中裝備欄位
  profileSheet.getRange(2, slotCol + 1).setValue(equipmentID);
  return { success: true };
}




function updateAdventureAttributesWithEquip() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName("Profile");
  const equipSheet = ss.getSheetByName("EquipmentMaster");
  const attrSheet = ss.getSheetByName("AdventureAttributes");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);
  const get = (key) => parseFloat(profile[key] || 0);

  const Intelligence = get("Intelligence"), IntelligenceLevel = get("IntelligenceLevel");
  const Physical = get("Physical"), PhysicalLevel = get("PhysicalLevel");
  const Creativity = get("Creativity"), CreativityLevel = get("CreativityLevel");
  const Sensitivity = get("Sensitivity"), SensitivityLevel = get("SensitivityLevel");
  const Grace = get("Grace"), GraceLevel = get("GraceLevel");
  const SelfDiscipline = get("SelfDiscipline");

  const Luck = 10 + SelfDiscipline * 0.5;
  const Agility = SensitivityLevel * 8 + Sensitivity * 0.4;
  const calc = (base, bonusField) => base + (get(bonusField) || 0);

  const Base = {
    BaseHP: PhysicalLevel * 80 + Physical * 0.5,
    BaseMP: IntelligenceLevel * 60 + Intelligence * 0.5,
    BasePhysicalAttack: PhysicalLevel * 10 + Physical * 0.3,
    BasePhysicalDefense: PhysicalLevel * 8 + SelfDiscipline * 0.5,
    BaseMagicAttack: IntelligenceLevel * 12 + Creativity * 0.5,
    BaseMagicDefense: IntelligenceLevel * 6 + Grace * 0.3 + Sensitivity * 0.3,
    BaseAgility: Agility,
    BaseAccuracy: GraceLevel * 6 + Intelligence * 0.2,
    BaseEvasion: Agility * 0.5 + Sensitivity * 0.3,
    BaseCriticalRate: Creativity * 0.3 + Luck * 0.5,
    BaseCriticalDamage: Creativity * 0.6 + Physical * 0.3,
    BaseLuck: Luck,
    BaseResistance: SelfDiscipline * 0.6 + Grace * 0.3,
  };

  const equippedSlots = ['Head', 'Body', 'Gloves', 'Belt', 'Shoes', 'Accessory', 'Weapon'];
  const equipIDs = equippedSlots.map(s => profile[s]).filter(Boolean);
  const equipData = equipSheet.getDataRange().getValues();
  const equipHeaders = equipData[0];
  const equipRows = equipData.slice(1).filter(row => equipIDs.includes(row[0]));

  const Equip = {};
  Object.keys(Base).forEach(k => {
    const attr = k.replace("Base", "");
    Equip["Equip" + attr] = 0;
  });

  equipRows.forEach(row => {
    Object.keys(Equip).forEach(k => {
      const field = k.replace("Equip", "") + "Boost"; // e.g. HPBoost
      const colIdx = equipHeaders.indexOf(field);
      if (colIdx !== -1) {
        Equip[k] += parseFloat(row[colIdx]) || 0;
      }
    });
  });

  const Final = {};
  Object.keys(Base).forEach(k => {
    const attr = k.replace("Base", "");
    Final["Final" + attr] =
      (Base[k] || 0) +
      (Equip["Equip" + attr] || 0) +
      (get(attr + "Bonus") || 0);
  });

  const finalData = {
    ...Base,
    ...Equip,
    ...Final
  };

  const attrHeaders = attrSheet.getRange(1, 1, 1, attrSheet.getLastColumn()).getValues()[0];
  const outputRow = attrHeaders.map(k => finalData[k] ?? '');

  attrSheet.getRange(2, 1, 1, outputRow.length).setValues([outputRow]);

  Logger.log("✅ 已更新 AdventureAttributes");
}

function getQuickStatus() {
  const ss = SpreadsheetApp.openById("16soa5MJP8XtRUsOvr6VjBArPSK0qZUeYs8pZnkwE1-s");
  const profileSheet = ss.getSheetByName("Profile");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  // TODO: 狀態系統尚未實作，未來從 StatusRules 表格評估狀態條件並產生 StatusList
  // 例如：髒兮兮 / 情緒低落 / 健康不佳 等
  const statusList = [];

  return {
    Coins: profile.Coins || 0,
    HonorPoints: profile.HonorPoints || 0,
    TodayActionPoints: profile.TodayActionPoints || 0,
    Cleanliness: profile.Cleanliness || 0,
    Mood: profile.Mood || 0,
    Energy: profile.Energy || 0,
    Health: profile.Health || 0,
    SelfDiscipline: profile.SelfDiscipline || 0,
    Intelligence: profile.Intelligence || 0,
    IntelligenceLevel: profile.IntelligenceLevel || 1,
    Physical: profile.Physical || 0,
    PhysicalLevel: profile.PhysicalLevel || 1,
    Creativity: profile.Creativity || 0,
    CreativityLevel: profile.CreativityLevel || 1,
    Sensitivity: profile.Sensitivity || 0,
    SensitivityLevel: profile.SensitivityLevel || 1,
    Grace: profile.Grace || 0,
    GraceLevel: profile.GraceLevel || 1,
    StatusList: statusList
  };
}
function getInventory() {
  const ss = SpreadsheetApp.getActive();
  const inventorySheet = ss.getSheetByName("Inventory");
  const itemSheet = ss.getSheetByName("ItemMaster");
  const equipSheet = ss.getSheetByName("EquipmentMaster");

  // 讀 Inventory
  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);

  // ItemMaster
  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemMap = {};
  itemData.slice(1).forEach(row => {
    const obj = {};
    itemHeaders.forEach((k, i) => obj[k] = row[i]);
    itemMap[obj.ItemID] = obj;
  });

  // EquipmentMaster
  const equipData = equipSheet.getDataRange().getValues();
  const equipHeaders = equipData[0];
  const equipMap = {};
  equipData.slice(1).forEach(row => {
    const obj = {};
    equipHeaders.forEach((k, i) => obj[k] = row[i]);
    equipMap[obj.EquipmentID] = obj;
  });

const result = invRows.map((row, idx) => {
  const inv = {};
  invHeaders.forEach((k, i) => inv[k] = row[i]);
  const id = inv.ItemID;
  const count = parseInt(inv.Count || 0);
  const equipped = inv.Equipped || false;

  // 👇 加入售價取得邏輯
  let sellPrice = 0;
  let honorSellPrice = 0;

  if (itemMap[id]) {
    sellPrice = parseInt(itemMap[id].SellPrice || 0);
    honorSellPrice = parseInt(itemMap[id].HonorSellPrice || 0);
  } else if (equipMap[id]) {
    sellPrice = parseInt(equipMap[id].SellPrice || 0);
    honorSellPrice = parseInt(equipMap[id].HonorSellPrice || 0);
  }

  let merged = {
    ItemID: id,
    Count: count,
    Equipped: equipped,
    SellPrice: sellPrice,
    HonorSellPrice: honorSellPrice,
    rowIndex: idx,
  };


    if (itemMap[id]) {
      // 🔹 來自 ItemMaster（消耗品、卷）
      Object.assign(merged, {
        ItemName: itemMap[id].ItemName || id,
        ItemType: itemMap[id].ItemType || "Item",
        Effect: itemMap[id].Effect || "",
        Description: itemMap[id].Description || "",
        IsVirtual: itemMap[id].IsVirtual || "Yes"
      });
    } else if (equipMap[id]) {
      // 🔹 來自 EquipmentMaster（裝備）
      Object.assign(merged, {
        ItemName: equipMap[id].Name || id,
        ItemType: "Equip",
        Effect: generateEffectFromEquip(equipMap[id]),
        Description: "裝備道具，加成能力",
        Slot: equipMap[id].Slot
      });
    } else {
      // ❌ 未知物品
      merged.ItemName = id;
      merged.ItemType = "Unknown";
      merged.Description = "❓ 無法辨識物品";
    }

    return merged;
  });

  return result;
}

function generateEffectFromEquip(row) {
  const boostFields = Object.keys(row).filter(k => k.endsWith("Boost") && row[k]);
  return boostFields.map(k => `${k}+${row[k]}`).join(",");
}
function getFieldMapping() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("FieldMapping");
  const values = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const zh = values[i][0];
    const en = values[i][1];
    if (zh && en) map[en] = zh;
  }

  return map;
}
function useItem(itemID) {
  const ss = SpreadsheetApp.getActive();
  const profileSheet = ss.getSheetByName("Profile");
  const inventorySheet = ss.getSheetByName("Inventory");
  const itemSheet = ss.getSheetByName("ItemMaster");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);
  const invIndex = invRows.findIndex(row => row[invHeaders.indexOf("ItemID")] === itemID);
  if (invIndex === -1) throw new Error("❌ 背包中找不到此道具");

  const countIdx = invHeaders.indexOf("Count");
  const count = parseInt(invRows[invIndex][countIdx]);
  if (count <= 0) throw new Error("⚠️ 數量不足");

  // 讀取效果
  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemRow = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === itemID);
  if (!itemRow) throw new Error("❌ 找不到道具資料");

  const effectStr = itemRow[itemHeaders.indexOf("Effect")] || '';
  const effectPairs = effectStr.split(',').map(p => p.trim()).filter(p => p.includes('+'));
  effectPairs.forEach(pair => {
    const [field, val] = pair.split('+');
    profile[field] = (parseFloat(profile[field]) || 0) + parseFloat(val);
  });

  // 更新 inventory：數量 -1 或刪除
  if (count - 1 <= 0) {
    inventorySheet.deleteRow(invIndex + 2);
  } else {
    invRows[invIndex][countIdx] = count - 1;
    inventorySheet.getRange(invIndex + 2, 1, 1, invHeaders.length).setValues([invRows[invIndex]]);
  }

  writeProfile(profile, "使用道具 - " + itemRow[itemHeaders.indexOf("ItemName")]);
}
/**
 * 販售道具／裝備（可指定數量）
 * @param {string} itemID 目標 ItemID / EquipmentID
 * @param {number} amount 使用者想賣出的數量 (>=1)
 * @return {string} 結果訊息
 */
function sellItem(itemID, amount) {
  amount = parseInt(amount, 10);
  if (isNaN(amount) || amount <= 0) throw new Error("⚠️ 請輸入正整數數量");

  const ss = SpreadsheetApp.getActive();
  const invS   = ss.getSheetByName("Inventory");
  const itemS  = ss.getSheetByName("ItemMaster");
  const equipS = ss.getSheetByName("EquipmentMaster");
  const profS  = ss.getSheetByName("Profile");

  /* -------- 讀取背包 -------- */
  const invData   = invS.getDataRange().getValues();
  const invHead   = invData[0];
  const invRows   = invData.slice(1);
  const idCol     = invHead.indexOf("ItemID");
  const cntCol    = invHead.indexOf("Count");
  const eqpCol    = invHead.indexOf("Equipped");

  const sameRows  = invRows
      .map((row, i) => ({ row, idx: i + 2 }))         // +2 → 實際試算表列號
      .filter(obj => obj.row[idCol] === itemID);

  if (!sameRows.length) throw new Error("❌ 找不到此物品");

  const totalCount   = sameRows.reduce((s, o) => s + (parseInt(o.row[cntCol]) || 0), 0);
  const equippedCnt  = sameRows.reduce((s, o) => s + ((o.row[eqpCol] === true || o.row[eqpCol] === "TRUE") ? (parseInt(o.row[cntCol]) || 0) : 0), 0);

  if (amount > totalCount) throw new Error("⚠️ 數量不足，背包只有 " + totalCount);

  // ▶ 若有裝備中，就必須留 1 件
  if (equippedCnt > 0 && totalCount - amount < 1) {
    throw new Error("⚠️ 目前正穿戴此裝備，至少要保留 1 件");
  }

  /* -------- 讀取售價 -------- */
  const equipRow = equipS.getDataRange().getValues()
      .slice(1)
      .find(r => r[equipS.getDataRange().getValues()[0].indexOf("EquipmentID")] === itemID);
  const itemRow  = itemS.getDataRange().getValues()
      .slice(1)
      .find(r => r[itemS.getDataRange().getValues()[0].indexOf("ItemID")] === itemID);

  let sellPrice = 0, honorSell = 0, itemName = itemID;
  if (equipRow) {
    const eh = equipS.getDataRange().getValues()[0];
    sellPrice   = parseInt(equipRow[eh.indexOf("SellPrice")])      || 0;
    honorSell   = parseInt(equipRow[eh.indexOf("HonorSellPrice")]) || 0;
    itemName    = equipRow[eh.indexOf("Name")]                     || itemID;
  } else if (itemRow) {
    const ih = itemS.getDataRange().getValues()[0];
    sellPrice   = parseInt(itemRow[ih.indexOf("SellPrice")])      || 0;
    honorSell   = parseInt(itemRow[ih.indexOf("HonorSellPrice")]) || 0;
    itemName    = itemRow[ih.indexOf("ItemName")]                 || itemID;
  } else {
    throw new Error("❓ 找不到物品資料");
  }
  if (sellPrice === 0 && honorSell === 0) throw new Error("❌ 此物品不可販售");

  /* -------- 扣背包數量 -------- */
  let remain = amount;                 // 還要扣幾件
  // 先從「未裝備」的列扣，最後才動到裝備列
  const sortFn = (a, b) => {
    const aEquipped = a.row[eqpCol] === true || a.row[eqpCol] === "TRUE";
    const bEquipped = b.row[eqpCol] === true || b.row[eqpCol] === "TRUE";
    return (aEquipped === bEquipped) ? 0 : (aEquipped ? 1 : -1);
  };
  sameRows.sort(sortFn);

  sameRows.forEach(obj => {
    if (remain <= 0) return;
    const rowCnt = parseInt(obj.row[cntCol]) || 0;
    const isEquipped = obj.row[eqpCol] === true || obj.row[eqpCol] === "TRUE";
    const minLeft = (isEquipped ? 1 : 0);              // 裝備列至少留 1

    const canRemove = Math.min(remain, rowCnt - minLeft);
    if (canRemove <= 0) return;

    const newCnt = rowCnt - canRemove;
    if (newCnt === 0) {
      invS.deleteRow(obj.idx);
    } else {
      invS.getRange(obj.idx, cntCol + 1).setValue(newCnt);
    }
    remain -= canRemove;
  });

  if (remain > 0) throw new Error("⚠️ 內部校正失敗，還剩 " + remain + " 未扣");

  /* -------- 加錢 / 加榮譽 -------- */
  const profH = profS.getRange(1,1,1,profS.getLastColumn()).getValues()[0];
  const profRow = profS.getRange(2,1,1,profH.length).getValues()[0];
  const profile = {}; profH.forEach((k,i)=>profile[k]=profRow[i]);

  if (sellPrice > 0) {
    profile.Coins = (parseInt(profile.Coins)||0) + sellPrice * amount;
    writeProfile(profile, `販售 ${itemName} x${amount} +${sellPrice*amount} 金幣`);
  } else {
    profile.HonorPoints = (parseInt(profile.HonorPoints)||0) + honorSell * amount;
    writeProfile(profile, `販售 ${itemName} x${amount} +${honorSell*amount} 榮譽點數`);
  }

  return `✅ 已販售 ${itemName} x${amount}`;
}
function unequip(slot) {
  const ss = SpreadsheetApp.getActive();
  const profileSheet = ss.getSheetByName("Profile");
  const inventorySheet = ss.getSheetByName("Inventory");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const slotCol = profileHeaders.indexOf(slot);
  if (slotCol === -1) throw new Error("❌ Profile 中找不到欄位: " + slot);

  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const oldEquipID = profileRow[slotCol];

  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const idCol = invHeaders.indexOf("ItemID");
  const equippedCol = invHeaders.indexOf("Equipped");

  for (let i = 1; i < invData.length; i++) {
    if (invData[i][idCol] === oldEquipID) {
      inventorySheet.getRange(i + 1, equippedCol + 1).setValue(false);
      break;
    }
  }

  profileSheet.getRange(2, slotCol + 1).setValue('');
  return { success: true };
}
function getAvailableEquipments(slot) {
  const ss = SpreadsheetApp.getActive();
  const invSheet = ss.getSheetByName("Inventory");
  const equipSheet = ss.getSheetByName("EquipmentMaster");

  const invData = invSheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);

  const idCol = invHeaders.indexOf("ItemID");
  const equippedCol = invHeaders.indexOf("Equipped");

  const equippedIDs = invRows
    .filter(r => String(r[equippedCol]).toLowerCase() !== 'true')
    .map(r => r[idCol]);

  const equipData = equipSheet.getDataRange().getValues();
  const equipHeaders = equipData[0];
  const slotCol = equipHeaders.indexOf("Slot");

  const candidates = equipData
    .slice(1)
    .filter(row =>
      equippedIDs.includes(row[0]) &&
      row[slotCol] === slot
    )
    .map(row => {
      const obj = {};
      equipHeaders.forEach((k, i) => obj[k] = row[i]);
      return obj;
    });

  Logger.log("📌 slot:", slot);
  Logger.log("📌 符合未裝備 ID:", equippedIDs);
  Logger.log("📌 選出候選裝備數量:", candidates.length);

  return { candidates };
}


