const SPREADSHEET_ID = '1OSkHqIGwq4xYEndtsrTk4Sc_EldHDeZIbvSg5L6djFs';

// ✅ 日期欄位一定要判斷是否為 Date 並轉為 "yyyy/MM/dd" 格式再進行比較
// 否則會導致 == 比對失敗、條件永遠不成立 顯示用途也要處理日期格式，避免出現 GMT/UTC 雜訊。
function doGet(e) {
  // ✅ 新增路由功能，區分遊戲主頁和核銷頁面
  if (e.parameter.page === 'verify' && e.parameter.token) {
    const template = HtmlService.createTemplateFromFile('verification');
    template.token = e.parameter.token; // 將 token 傳給 HTML 樣板
    return template.evaluate().setTitle('遊戲獎勵核銷').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ✅ 新增：提供一個頁面來查看所有待核銷的項目
  if (e.parameter.page === 'admin') {
    const template = HtmlService.createTemplateFromFile('admin');
    return template.evaluate().setTitle('管理後台 - 待核銷列表').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 預設回傳遊戲主頁
  return HtmlService.createHtmlOutputFromFile('index').setTitle('我的遊戲').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getProfileData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Profile');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // --- 1. 檢查是否需要執行每日結算 (方便測試) ---
  let profileRow = sheet.getRange(2, 1, 1, headers.length).getValues()[0];
  let initialProfile = {};
  headers.forEach((key, i) => initialProfile[key] = profileRow[i]);

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const lastUpdateDate = initialProfile.LastUpdateDate;
  const lastUpdateStr = (lastUpdateDate instanceof Date) 
    ? Utilities.formatDate(lastUpdateDate, Session.getScriptTimeZone(), "yyyy/MM/dd") 
    : "";

  if (lastUpdateStr !== todayStr) {
    Logger.log(`[getProfileData] 偵測到今天尚未結算 (上次結算: ${lastUpdateStr})，正在執行 applyEndOfDayUpdates()...`);
    applyEndOfDayUpdates(); // 執行每日結算
    // 結算後，必須重新讀取資料
    profileRow = sheet.getRange(2, 1, 1, headers.length).getValues()[0]; 
    Logger.log(`[getProfileData] 每日結算執行完畢，已重新讀取資料。`);
  }

  // --- 2. 處理並回傳最新的玩家資料 ---
  const profile = {};
  headers.forEach((key, i) => profile[key] = profileRow[i]);

  const birthdayFormatted = (profile.birthday instanceof Date)
    ? Utilities.formatDate(profile.birthday, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : (profile.birthday || '');

  const lastSurveyDateFormatted = (profile.LastSurveyDate instanceof Date)
    ? Utilities.formatDate(profile.LastSurveyDate, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : (profile.LastSurveyDate || '');
  const surveyFilledToday = (lastSurveyDateFormatted === todayStr);

  // --- 3. 狀態圖片更換邏輯 ---
  const activeStatuses = evaluateStatusRules(profile);
  Logger.log(`[getProfileData] 評估出的生效狀態: ${JSON.stringify(activeStatuses)}`); // ✅ 新增偵錯日誌
  const overrideStatus = activeStatuses
    .filter(s => s.CharacterOverrideFile)
    .sort((a, b) => (a.Priority || 999) - (b.Priority || 999))[0];

  let finalCharacterFile = profile.currentCharacter;
  if (overrideStatus) {
    finalCharacterFile = overrideStatus.CharacterOverrideFile;
    Logger.log(`狀態圖片覆蓋：因 [${overrideStatus.狀態名稱}]，角色圖片更換為 ${finalCharacterFile}`);
  }

  // ✅ 新增：將狀態列表也一併回傳
  const statusList = activeStatuses.map(status => ({
    StatusName: status.狀態名稱,
    Effect: status.效果說明
  }));

  return {
    playerName: profile.PlayerName,
    birthday: birthdayFormatted,
    coins: profile.Coins,
    honorPoints: profile.HonorPoints,
    cleanliness: profile.Cleanliness,
    mood: profile.Mood,
    energy: profile.Energy,
    health: profile.Health,
    selfDiscipline: profile.SelfDiscipline, 
    backgroundUrl: "https://renewmiffy.github.io/mygame/img/bg/" + profile.currentBackground,
    characterUrl: "https://renewmiffy.github.io/mygame/img/char/" + finalCharacterFile,
    surveyFilledToday: surveyFilledToday,
    debugLog: activeStatuses.debugLog, // ✅ 將偵錯日誌一起回傳
    StatusList: statusList // ✅ 新增
  };
}
function getSurveyQuestions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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

  profile.Cleanliness = (parseFloat(profile.Cleanliness) || 0) - 20;
  profile.LastSurveyDate = today;

  writeProfile(profile, "問卷");

  const logRow = [today, ...surveyAnswers, new Date()];
  logSheet.appendRow(logRow);
}

function getRecentLogs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  };

  Object.entries(limits).forEach(([field, [min, max]]) => {
    const val = parseFloat(profile[field] ?? 0);
    profile[field] = Math.min(max, Math.max(min, isNaN(val) ? 0 : val));
  });
}

function writeProfile(profile, source = "系統") {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
    "Coins", "HonorPoints"
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const missionSheet = ss.getSheetByName("MissionCenter");
  const profileSheet = ss.getSheetByName("Profile");
  const doneTasksSheet = ss.getSheetByName("DailyTasks");
  const skillSheet = ss.getSheetByName("SkillMaster"); // ✅ 1. 讀取技能表

  const missionData = missionSheet.getDataRange().getValues();
  const missionHeaders = missionData[0];
  const missionRows = missionData.slice(1);

  const getMissionField = (row, field) => {
    const idx = missionHeaders.indexOf(field);
    if (idx === -1) throw new Error(`❌ 找不到欄位：「${field}」`);
    return row[idx];
  };

  // ✅ 2. 建立技能資料的快取 Map，方便快速查找
  const skillData = skillSheet.getDataRange().getValues();
  const skillHeaders = skillData[0];
  const skillMap = {};
  skillData.slice(1).forEach(row => {
    const skillId = row[skillHeaders.indexOf("SkillID")];
    if (skillId) {
      skillMap[skillId] = {
        totalDone: parseInt(row[skillHeaders.indexOf("TotalDoneCount")] || 0),
        streak: parseInt(row[skillHeaders.indexOf("StreakCount")] || 0)
      };
    }
  });

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
    let currentProgress = 0; // ✅ 3. 新增變數
    let targetValue = 1;     // ✅ 3. 新增變數

    if (conditionType === "DailyTaskDoneCount") {
      currentProgress = totalTaskDone;
      targetValue = parseInt(param);
      fulfilled = currentProgress >= targetValue;
    } else if (conditionType === "TaskDoneToday") {
      fulfilled = todayTaskIds.includes(param);
      currentProgress = fulfilled ? 1 : 0;
      targetValue = 1;
    } else if (conditionType === "TaskDoneCount") {
      const [taskId, count] = param.split(":");
      const taskRow = dailyData.find(r => r[0] === taskId);
      currentProgress = taskRow ? parseInt(taskRow[5] || 0) : 0;
      targetValue = parseInt(count);
      fulfilled = currentProgress >= targetValue;
    } else if (conditionType === "TotalDoneCount") { // ✅ 變更：直接使用欄位名稱，更直觀
      const [skillId, count] = param.split(":");
      targetValue = parseInt(count);
      const skillInfo = skillMap[skillId];
      if (skillInfo) {
        currentProgress = skillInfo.totalDone;
      } else {
        currentProgress = 0; // 找不到技能，進度為 0
      }
      fulfilled = currentProgress >= targetValue;
    } else if (conditionType === "StreakCount") { // ✅ 變更：直接使用欄位名稱，更直觀
      const [skillId, count] = param.split(":");
      targetValue = parseInt(count);
      const skillInfo = skillMap[skillId];
      if (skillInfo) {
        currentProgress = skillInfo.streak;
      } else {
        currentProgress = 0;
      }
      fulfilled = currentProgress >= targetValue;
    }

    let claimed = false;
    if (type === "Daily") {
      claimed = formatYMD(lastClaimed) === today;
    } else {
      claimed = !!lastClaimed && !repeatable;
    }

    let rewardText = "";
    const rewardJsonString = getMissionField(row, "獎勵") || "{}";
    try {
      const rewardObj = JSON.parse(rewardJsonString);
      // 檢查確保解析出來的是一個物件
      if (typeof rewardObj === 'object' && rewardObj !== null && !Array.isArray(rewardObj)) {
        rewardText = Object.entries(rewardObj)
          .map(([key, val]) => `+${val} ${key}`)
          .join(" / ");
      } else {
        rewardText = rewardJsonString; // 如果不是物件，直接顯示原始文字
      }
    } catch (e) {
      rewardText = `${rewardJsonString} (❌ 獎勵格式錯誤)`;
    }

    return {
      id: missionId,
      name: name,
      description: name, // 直接使用任務名稱，避免前端顯示重複
      rewardText: rewardText,
      fulfilled: fulfilled,
      claimed: claimed,
      displayOrder: displayOrder,
      rowIndex: i,
      type: type, // 加入類型給一鍵領取判斷用
      currentProgress: currentProgress, // ✅ 5. 回傳進度
      targetValue: targetValue        // ✅ 5. 回傳目標
    };
  });

  return result.sort((a, b) => a.displayOrder - b.displayOrder);
}


// ✅ claimDailyTask()：支援 JSON 獎勵 + LastClaimedDate 寫入
function claimDailyTask(missionId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("SkillMaster");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  
  return rows.map(row => {
    const skill = {};
    let lastDateStr = "";
    headers.forEach((h, i) => {
      if (h === "LastDoneDate" && row[i] instanceof Date) {
        lastDateStr = Utilities.formatDate(row[i], Session.getScriptTimeZone(), "yyyy/MM/dd");
        skill[h] = lastDateStr;
      } else {
        skill[h] = row[i];
      }
    });
    skill.learnedToday = (lastDateStr === todayStr);
    return skill;
  });
}
function getAttributeMap() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("FieldMapping");
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


/**
 * [核心狀態評估函式] 根據玩家屬性，回傳所有觸發的狀態規則。
 * @param {object} profile - 玩家的 profile 物件。
 * @returns {Array<object>} - 一個包含所有被觸發的狀態規則物件的陣列。
 */
function evaluateStatusRules(profile) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const statusSheet = ss.getSheetByName("StatusRules");
  if (!statusSheet) {
    Logger.log("⚠️ StatusRules 工作表不存在。");
    return [];
  }

  const data = statusSheet.getDataRange().getValues();
  const headers = data[0].map(h => h.trim()); // ✅ 清理標頭前後的空白
  const rules = data.slice(1);
  const activeStatuses = [];
  const debugLog = []; // ✅ 新增一個陣列來收集日誌

  Logger.log("--- 開始評估狀態規則 ---");

  rules.forEach((ruleRow, index) => {
    const rule = {};
    headers.forEach((h, i) => rule[h] = ruleRow[i]);

    const statusName = rule["狀態名稱"] || `規則 #${index + 1}`;
    const triggerField = rule["觸發屬性"];
    const comparison = rule["比較符號"];
    
    if (!triggerField || !comparison) {
      debugLog.push(`- [${statusName}] 跳過：觸發屬性或比較符號為空。`);
      return; // Skip this rule if essential data is missing
    }

    const playerValue = parseFloat(profile[triggerField]);
    const threshold = parseFloat(rule["閾值"]);

    debugLog.push(`- 正在檢查 [${statusName}]：屬性=${triggerField}, 條件=${comparison} ${threshold}`);

    if (isNaN(playerValue) || isNaN(threshold)) {
      debugLog.push(`  - ↳ 跳過：玩家數值 (${profile[triggerField]}) 或閾值 (${rule["閾值"]}) 不是有效的數字。`);
      return;
    }
    
    debugLog.push(`  - ↳ 比較：玩家的 ${triggerField} (${playerValue}) ${comparison} ${threshold} ?`);

    let isTriggered = false;
    if (comparison === '<' && playerValue < threshold) isTriggered = true;
    else if (comparison === '<=' && playerValue <= threshold) isTriggered = true;
    else if (comparison === '>' && playerValue > threshold) isTriggered = true;
    else if (comparison === '>=' && playerValue >= threshold) isTriggered = true;
    else if (comparison === '==' && playerValue == threshold) isTriggered = true;

    if (isTriggered) {
      debugLog.push(`  - ↳ ✅ 結果：觸發！`);
      activeStatuses.push({
        狀態名稱: statusName,
        屬性影響JSON: rule["屬性影響JSON"],
        效果說明: rule["效果說明"],
        Priority: rule["Priority"], // ✅ 新增
        CharacterOverrideFile: rule["CharacterOverrideFile"] // ✅ 新增
      });
    } else {
      debugLog.push(`  - ↳ ❌ 結果：未觸發。`);
    }
  });

  debugLog.push(`--- 狀態評估結束，共觸發 ${activeStatuses.length} 個狀態 ---`);
  // 將結果和日誌一起回傳
  activeStatuses.debugLog = debugLog;
  return activeStatuses; 
}

function getQuickStatus() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);
 
  // ✅ 呼叫新的狀態評估函式
  const activeStatuses = evaluateStatusRules(profile);
  const statusList = activeStatuses.map(status => ({
    StatusName: status.狀態名稱,
    Effect: status.效果說明
  }));

  // ✅ 新增：狀態圖片更換邏輯，與 getProfileData() 同步
  const overrideStatus = activeStatuses
    .filter(s => s.CharacterOverrideFile)
    .sort((a, b) => (a.Priority || 999) - (b.Priority || 999))[0];

  let finalCharacterFile = profile.currentCharacter;
  if (overrideStatus) {
    finalCharacterFile = overrideStatus.CharacterOverrideFile;
  }

  return {
    Coins: profile.Coins || 0,
    HonorPoints: profile.HonorPoints || 0,
    Cleanliness: profile.Cleanliness || 0,
    Mood: profile.Mood || 0,
    Energy: profile.Energy || 0,
    Health: profile.Health || 0,
    SelfDiscipline: profile.SelfDiscipline || 0,
    StatusList: statusList,
    // ✅ 新增：回傳計算後的角色圖片 URL
    characterUrl: "https://renewmiffy.github.io/mygame/img/char/" + finalCharacterFile
  };
}
function getInventory() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const inventorySheet = ss.getSheetByName("Inventory");
  const itemSheet = ss.getSheetByName("ItemMaster");

  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);

  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemMap = {};
  itemData.slice(1).forEach(row => {
    const obj = {};
    itemHeaders.forEach((k, i) => obj[k] = row[i]);
    itemMap[obj.ItemID] = obj;
  });

  return invRows.map((row, idx) => {
    const inv = {};
    invHeaders.forEach((k, i) => inv[k] = row[i]);
    const id = inv.ItemID;
    const count = parseInt(inv.Count || 0);

    const base = {
      ItemID: id,
      Count: count,
      rowIndex: idx,
      Equipped: false,
      SellPrice: 0,
      HonorSellPrice: 0,
    };

    const ref = itemMap[id];
    if (ref) {
      Object.assign(base, {
        ItemName: ref.ItemName || id,
        ItemType: ref.ItemType || "Consumable",
        Description: ref.Description || "",
        Effect: ref.Effect || "",
        IsSellable: ref.IsSellable,
        SellPrice: parseInt(ref.SellPrice || 0),
        HonorSellPrice: parseInt(ref.HonorSellPrice || 0),
        Rarity: ref.Rarity || 0
      });
    } else {
      base.ItemName = id;
      base.ItemType = "Unknown";
      base.Description = "❓ 無資料";
    }

    return base;
  });
}


function getFieldMapping() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("FieldMapping");
  const values = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const zh = values[i][0];
    const en = values[i][1];
    if (zh && en) map[en] = zh;
  }

  return map;
}
function useItem(itemID, quantity) {
  // ✅ 輔助函式：執行單次兩階段抽獎
  function _performSinglePull(lootTable, allItemData, allItemHeaders) {
      const totalWeight = lootTable.reduce((sum, item) => sum + (item.Weight || 0), 0);
      if (totalWeight <= 0) throw new Error(`❌ 寶箱的總權重為 0。`);
  
      let random = Math.random() * totalWeight;
      let wonLoot = null;
  
      for (const loot of lootTable) {
        random -= (loot.Weight || 0);
        if (random <= 0) {
          wonLoot = loot;
          break;
        }
      }
      if (wonLoot === null) wonLoot = lootTable[lootTable.length - 1];
  
      const type = wonLoot.Type || "ItemRarity"; // 預設為舊格式
  
      if (type === "Currency") {
        return { type: 'currency', rewardID: wonLoot.RewardID, quantity: wonLoot.Quantity, displayName: wonLoot.DisplayName, rarity: 'currency' };
      } else { // type === "ItemRarity"
        const selectedRarity = wonLoot.Rarity;
        const rarityCol = allItemHeaders.indexOf("Rarity");
        const potentialItems = allItemData.slice(1).filter(row => String(row[rarityCol]) === String(selectedRarity));
        if (potentialItems.length === 0) throw new Error(`❌ 在 ItemMaster 中找不到任何 Rarity 為 [${selectedRarity}] 的道具可供抽取。`);
        const wonItemRow = potentialItems[Math.floor(Math.random() * potentialItems.length)];
        const wonItemID = wonItemRow[allItemHeaders.indexOf("ItemID")];
        const wonItemName = wonItemRow[allItemHeaders.indexOf("ItemName")] || wonItemID;
        return { type: 'item', itemID: wonItemID, itemName: wonItemName, rarity: selectedRarity };
      }
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");
  const inventorySheet = ss.getSheetByName("Inventory");
  const itemSheet = ss.getSheetByName("ItemMaster");
  const logSheet = ss.getSheetByName("PlayerLog");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);
  const invIndex = invRows.findIndex(row => row[invHeaders.indexOf("ItemID")] === itemID);
  if (invIndex === -1) throw new Error("❌ 背包中找不到此道具");

  quantity = parseInt(quantity) || 1;
  if (quantity <= 0) throw new Error("⚠️ 使用數量必須大於 0");

  const countIdx = invHeaders.indexOf("Count");
  const count = parseInt(invRows[invIndex][countIdx]);
  if (count < quantity) throw new Error(`⚠️ 數量不足，需要 ${quantity} 個，但你只有 ${count} 個。`);

  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemRow = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === itemID);
  if (!itemRow) throw new Error("❌ 找不到道具主資料 (ItemMaster)");

  const itemMaster = {};
  itemHeaders.forEach((k, i) => itemMaster[k] = itemRow[i]);
  const itemName = itemMaster.ItemName || itemID;
  const itemType = itemMaster.ItemType || 'Consumable'; // 預設為消耗品

  // 對於一次性使用的道具，強制數量為 1
  const singleUseTypes = ['TreasureChest', 'Redeemable'];
  if (singleUseTypes.includes(itemType) && quantity > 1) {
    quantity = 1; // 強制使用 1 個
  }

  // ----------------------------------------------------------------
  // 1. 處理寶箱 (TreasureChest)
  // ----------------------------------------------------------------
  if (itemType === 'TreasureChest') {
    const lootTableJson = itemMaster.LootTableJSON || '[]';
    let lootTable;
    try {
      lootTable = JSON.parse(lootTableJson);
      if (!Array.isArray(lootTable) || lootTable.length === 0) {
        throw new Error("獎池為空或格式不正確。");
      }
    } catch (e) {
      throw new Error(`❌ 無法解析寶箱 [${itemName}] 的獎池設定 (LootTableJSON): ${e.message}`);
    }

    // ✅ 使用重構後的輔助函式執行單抽
    const pullResult = _performSinglePull(lootTable, itemData, itemHeaders);

    // ✅ 修正：將消耗寶箱的邏輯移到最前面，確保一定會執行
    // 消耗寶箱
    if (count - quantity <= 0) {
      inventorySheet.deleteRow(invIndex + 2);
    } else {
      inventorySheet.getRange(invIndex + 2, countIdx + 1).setValue(count - quantity);
    }

    // --- 發放獎勵 ---
    let resultMessage = "";
    if (pullResult.type === 'currency') {
      profile[pullResult.rewardID] = (parseFloat(profile[pullResult.rewardID]) || 0) + parseFloat(pullResult.quantity);
      writeProfile(profile, `開啟寶箱 - ${itemName}`);
      resultMessage = `恭喜！你從 ${itemName} 中獲得了 ${pullResult.displayName} (+${pullResult.quantity})！`;
    } else {
      // 發放道具
      const invData = inventorySheet.getDataRange().getValues();
      const invHeaders = invData[0];
      const invRows = invData.slice(1);
      const itemIDCol = invHeaders.indexOf("ItemID");
      const countCol = invHeaders.indexOf("Count");
      const existingItemIndex = invRows.findIndex(r => r[itemIDCol] === pullResult.itemID);

      if (existingItemIndex !== -1) {
        const sheetRowIndex = existingItemIndex + 2;
        const currentCount = parseInt(invRows[existingItemIndex][countCol]) || 0;
        inventorySheet.getRange(sheetRowIndex, countCol + 1).setValue(currentCount + 1);
      } else {
        const newRow = invHeaders.map(h => (h === "ItemID") ? pullResult.itemID : (h === "Count" ? 1 : ""));
        inventorySheet.appendRow(newRow);
      }
      resultMessage = `恭喜！你從 ${itemName} 中獲得了【${pullResult.rarity}★】${pullResult.itemName}！`;
    }

    // 寫入日誌
    const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    const logEntry = {
        CreatedAt: new Date(),
        Type: 'action',
        ActionName: `開啟寶箱`,
        Source: resultMessage.replace('恭喜！', '').trim()
    };
    const logRow = logHeaders.map(header => logEntry[header] || '');
    logSheet.appendRow(logRow);
    return resultMessage;
  }
  // ----------------------------------------------------------------
  // ✅ 新增：處理固定內容禮包 (FixedBundle)
  // ----------------------------------------------------------------
  else if (itemType === 'FixedBundle') {
    const effectJson = itemMaster.Effect || '{}';
    const itemsGranted = [];
    const itemsToAdd = {};

    try {
      const effects = JSON.parse(effectJson);
      if (effects.Items && Array.isArray(effects.Items)) {
        effects.Items.forEach(itemToGrant => {
          if (itemToGrant.ItemID && itemToGrant.Quantity > 0) {
            itemsToAdd[itemToGrant.ItemID] = (itemsToAdd[itemToGrant.ItemID] || 0) + (parseInt(itemToGrant.Quantity) * quantity);
          }
        });
      } else {
        throw new Error("Effect JSON 中缺少有效的 'Items' 陣列。");
      }
    } catch (e) {
      throw new Error(`❌ 無法解析禮包 [${itemName}] 的效果設定 (Effect): ${e.message}`);
    }

    if (Object.keys(itemsToAdd).length === 0) {
      throw new Error(`❌ 禮包 [${itemName}] 的內容為空。`);
    }

    // --- 批次發放道具 (此處假設禮包內道具皆可堆疊) ---
    const invData = inventorySheet.getDataRange().getValues();
    const invHeaders = invData[0];
    const invRows = invData.slice(1);
    const itemIDCol = invHeaders.indexOf("ItemID");
    const countCol = invHeaders.indexOf("Count");

    const inventoryMap = {};
    invRows.forEach((row, index) => {
      inventoryMap[row[itemIDCol]] = { count: parseInt(row[countCol]) || 0, sheetRow: index + 2 };
    });

    const rowsToAppend = [];
    Object.entries(itemsToAdd).forEach(([id, qty]) => {
      if (inventoryMap[id]) {
        inventorySheet.getRange(inventoryMap[id].sheetRow, countCol + 1).setValue(inventoryMap[id].count + qty);
      } else {
        rowsToAppend.push(invHeaders.map(h => (h === "ItemID") ? id : (h === "Count" ? qty : "")));
      }
      const grantedItemName = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === id)?.[itemHeaders.indexOf("ItemName")] || id;
      itemsGranted.push(`${grantedItemName} x${qty}`);
    });

    if (rowsToAppend.length > 0) {
      inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, rowsToAppend.length, invHeaders.length).setValues(rowsToAppend);
    }

    // 消耗禮包
    if (count - quantity <= 0) { inventorySheet.deleteRow(invIndex + 2); } 
    else { inventorySheet.getRange(invIndex + 2, countIdx + 1).setValue(count - quantity); }
    
    writeProfile(profile, `開啟禮包 - ${itemName}`);
    return `✅ 已開啟 ${itemName} x${quantity}！獲得：${itemsGranted.join(', ')}`;
  }
  // ----------------------------------------------------------------
  // ✅ 新增：處理貨幣包 (CurrencyPouch)
  // ----------------------------------------------------------------
  else if (itemType === 'CurrencyPouch') {
    const effectJson = itemMaster.Effect || '{}';
    let effectApplied = false;
    const rewardsMessage = [];

    try {
      const effects = JSON.parse(effectJson);
      Object.entries(effects).forEach(([field, value]) => {
        if (profile.hasOwnProperty(field)) {
          const gain = parseFloat(value) * quantity;
          profile[field] = (parseFloat(profile[field]) || 0) + gain;
          effectApplied = true;
          rewardsMessage.push(`${field} +${gain}`);
        }
      });
    } catch (e) {
      throw new Error(`❌ 無法解析貨幣包 [${itemName}] 的效果設定 (Effect): ${e.message}`);
    }

    if (!effectApplied) { throw new Error(`❌ 貨幣包 [${itemName}] 的效果設定無效。`); }

    if (count - quantity <= 0) { inventorySheet.deleteRow(invIndex + 2); } 
    else { inventorySheet.getRange(invIndex + 2, countIdx + 1).setValue(count - quantity); }
    
    writeProfile(profile, `開啟貨幣包 - ${itemName}`);
    return `✅ 已開啟 ${itemName} x${quantity}！獲得：${rewardsMessage.join(', ')}`;
  }
  // ----------------------------------------------------------------
  // 2. 處理兌換券 (Redeemable)
  // ----------------------------------------------------------------
  else if (itemType === 'Redeemable') {
    // 這是兌換券，觸發核銷流程
    const token = Utilities.getUuid();
    const redemptionSheet = ss.getSheetByName('RedemptionLog') || ss.insertSheet('RedemptionLog');
    if (redemptionSheet.getLastRow() === 0) {
      redemptionSheet.appendRow(['Token', 'ItemID', 'ItemName', 'Status', 'RequestDate', 'ProcessDate', 'PlayerName']);
    }
    redemptionSheet.appendRow([token, itemID, itemName, 'Pending', new Date(), '', profile.PlayerName]);

    const adminUrl = ScriptApp.getService().getUrl() + '?page=admin';
    const url = ScriptApp.getService().getUrl() + '?page=verify&token=' + token;
    const subject = `[遊戲獎勵兌換] ${profile.PlayerName} 請求兌換：${itemName}`;
    const body = `
      <h3>您好！</h3>
      <p>玩家 <strong>${profile.PlayerName}</strong> 在遊戲中請求兌換以下實體獎勵：</p>
      <p style="font-size: 18px; font-weight: bold;">獎勵名稱： ${itemName}</p>
      <p>請點擊以下連結，單獨處理此項兌換：</p>
      <p><a href="${url}" style="font-size: 16px; padding: 10px 15px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">處理「${itemName}」</a></p>
      <hr style="margin: 20px 0;">
      <p style="font-size: 14px;">或者，您可以點擊下方連結，一次查看所有待處理的項目：</p>
      <p><a href="${adminUrl}" style="font-size: 16px; padding: 8px 12px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px;">查看所有待核銷列表</a></p>
    `;
    MailApp.sendEmail({ to: 'renewmiffy@gmail.com', subject: subject, htmlBody: body });

  } 
  // ----------------------------------------------------------------
  // 3. 處理一般消耗品 (Consumable)
  // ----------------------------------------------------------------
  else {
    // 這是一般消耗品，套用效果
    const effectJson = itemMaster.Effect || '{}';
    try {
      const effects = JSON.parse(effectJson);
      Object.entries(effects).forEach(([field, value]) => {
        if (profile.hasOwnProperty(field)) {
          profile[field] = (parseFloat(profile[field]) || 0) + (parseFloat(value) * quantity);
        }
      });
    } catch (e) { /* 忽略錯誤 */ }
  }

  if (count - quantity <= 0) {
    inventorySheet.deleteRow(invIndex + 2);
  } else {
    inventorySheet.getRange(invIndex + 2, countIdx + 1).setValue(count - quantity);
  }

  if (itemType === 'Redeemable') {
    // ✅【修正】直接寫入單筆日誌，而不是呼叫會清空屬性的 writeProfile({}, ...)
    const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    const logEntry = {
        CreatedAt: new Date(),
        Type: 'action',
        ActionName: `請求兌換 - ${itemName}`,
        Source: `請求兌換 - ${itemName}`
    };
    const logRow = logHeaders.map(header => logEntry[header] || '');
    logSheet.appendRow(logRow);
    return "✅ 兌換請求已發送！請等待家長為您核准。";
  } else {
    writeProfile(profile, `使用道具 - ${itemName}`);
    return `✅ 已使用 ${itemName} x${quantity}！`;
  }
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

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) throw new Error("❌ 無法開啟指定的試算表。");

  const invS   = ss.getSheetByName("Inventory");
  const itemS  = ss.getSheetByName("ItemMaster");
  const profS  = ss.getSheetByName("Profile");

  if (!invS) throw new Error("❌ 找不到 'Inventory' 工作表。");
  if (!itemS) throw new Error("❌ 找不到 'ItemMaster' 工作表。");
  if (!profS) throw new Error("❌ 找不到 'Profile' 工作表。");

  /* -------- 讀取背包 -------- */
  const invData   = invS.getDataRange().getValues();
  const invHead   = invData[0];
  const invRows   = invData.slice(1);
  const idCol     = invHead.indexOf("ItemID");
  const cntCol    = invHead.indexOf("Count");

  const sameRows  = invRows
      .map((row, i) => ({ row, idx: i + 2 }))         // +2 → 實際試算表列號
      .filter(obj => obj.row[idCol] === itemID);

  if (!sameRows.length) throw new Error("❌ 找不到此物品");

  const totalCount = sameRows.reduce((s, o) => s + (parseInt(o.row[cntCol]) || 0), 0);

  if (amount > totalCount) throw new Error("⚠️ 數量不足，背包只有 " + totalCount);

  /* -------- 讀取售價 -------- */
  const itemRow  = itemS.getDataRange().getValues()
      .slice(1)
      .find(r => r[itemS.getDataRange().getValues()[0].indexOf("ItemID")] === itemID);

  let sellPrice = 0, honorSell = 0, itemName = itemID;
  if (itemRow) {
    const ih = itemS.getDataRange().getValues()[0];
    sellPrice   = parseInt(itemRow[ih.indexOf("SellPrice")])      || 0;
    honorSell   = parseInt(itemRow[ih.indexOf("HonorSellPrice")]) || 0;
    itemName    = itemRow[ih.indexOf("ItemName")]                 || itemID;
  } else {
    throw new Error("❓ 找不到物品資料 (ItemMaster)");
  }
  if (sellPrice === 0 && honorSell === 0) throw new Error("❌ 此物品不可販售");

  /* -------- 扣背包數量 -------- */
  let remainingToRemove = amount;
  const rowsToDelete = [];
  const rowsToUpdate = [];

  for (const obj of sameRows) {
    if (remainingToRemove <= 0) break;
    const rowCnt = parseInt(obj.row[cntCol]) || 0;
    const amountToRemoveFromThisRow = Math.min(remainingToRemove, rowCnt);

    if (amountToRemoveFromThisRow > 0) {
      const newCount = rowCnt - amountToRemoveFromThisRow;
      if (newCount === 0) {
        rowsToDelete.push(obj.idx);
      } else {
        rowsToUpdate.push({idx: obj.idx, count: newCount});
      }
      remainingToRemove -= amountToRemoveFromThisRow;
    }
  }

  // 執行更新
  rowsToUpdate.forEach(update => {
    invS.getRange(update.idx, cntCol + 1).setValue(update.count);
  });

  // 從後往前刪除，避免 index 錯亂
  rowsToDelete.sort((a, b) => b - a).forEach(idx => {
    invS.deleteRow(idx);
  });

  if (remainingToRemove > 0) throw new Error("⚠️ 內部校正失敗，還剩 " + remainingToRemove + " 未扣");

  /* -------- 加錢 / 加榮譽 -------- */
  const profH = profS.getRange(1,1,1,profS.getLastColumn()).getValues()[0];
  const profRow = profS.getRange(2,1,1,profH.length).getValues()[0];
  const profile = {}; profH.forEach((k,i)=>profile[k]=profRow[i]);

  if (sellPrice > 0) {
    profile.Coins = (parseInt(profile.Coins)||0) + sellPrice * amount;
    writeProfile(profile, `販售 ${itemName} x${amount}`);
  } else {
    profile.HonorPoints = (parseInt(profile.HonorPoints)||0) + honorSell * amount;
    writeProfile(profile, `販售 ${itemName} x${amount}`);
  }

  return `✅ 已販售 ${itemName} x${amount}`;
}

/**
 * 取得商店中所有可購買的商品列表
 * @returns {Array<object>} - 商品物件的陣列
 */
function getShopItems() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const itemSheet = ss.getSheetByName("ItemMaster");
  if (!itemSheet) throw new Error("❌ 找不到 'ItemMaster' 工作表。");

  const data = itemSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const purchasableCol = headers.indexOf("IsPurchasable");
  if (purchasableCol === -1) {
    Logger.log("ItemMaster 中沒有 IsPurchasable 欄位，無法提供商品。");
    return [];
  }

  return rows
    .filter(row => row[purchasableCol] === true)
    .map(row => {
      const item = {};
      headers.forEach((h, i) => item[h] = row[i]);
      return {
        ItemID: item.ItemID,
        ItemName: item.ItemName,
        ItemType: item.ItemType,
        Description: item.Description,
        Effect: item.Effect,
        BuyPrice: parseInt(item.BuyPrice || 0),
        HonorBuyPrice: parseInt(item.HonorBuyPrice || 0)
      };
    });
}

/**
 * 處理玩家購買商品的邏輯
 * @param {string} itemID - 欲購買的商品 ID
 * @returns {string} - 執行結果訊息
 */
function buyItem(itemID) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");
  const itemSheet = ss.getSheetByName("ItemMaster");
  const inventorySheet = ss.getSheetByName("Inventory");

  if (!profileSheet || !itemSheet || !inventorySheet) throw new Error("❌ 找不到必要的資料表 (Profile/ItemMaster/Inventory)。");

  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemRow = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === itemID);
  if (!itemRow) throw new Error("❌ 找不到此商品。");

  const item = {};
  itemHeaders.forEach((h, i) => item[h] = itemRow[i]);
  if (item.IsPurchasable !== true) throw new Error("❌ 此商品不可購買。");

  const price = parseInt(item.BuyPrice || 0);
  const honorPrice = parseInt(item.HonorBuyPrice || 0);

  // ✅ 修正：直接讀取 Profile 工作表，避免使用 getProfileData() 導致鍵值大小寫不符和資料遺失
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  if (price > 0 && (parseInt(profile.Coins) || 0) < price) throw new Error(`⚠️ 金幣不足！需要 ${price}。`);
  if (honorPrice > 0 && (parseInt(profile.HonorPoints) || 0) < honorPrice) throw new Error(`⚠️ 榮譽點數不足！需要 ${honorPrice}。`);

  if (price > 0) profile.Coins = (parseInt(profile.Coins) || 0) - price;
  if (honorPrice > 0) profile.HonorPoints = (parseInt(profile.HonorPoints) || 0) - honorPrice;

  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);
  const itemIDCol = invHeaders.indexOf("ItemID");
  const countCol = invHeaders.indexOf("Count");
  const existingItemIndex = invRows.findIndex(r => r[itemIDCol] === itemID);

  if (existingItemIndex !== -1) {
    const sheetRowIndex = existingItemIndex + 2;
    const currentCount = parseInt(invRows[existingItemIndex][countCol]) || 0;
    inventorySheet.getRange(sheetRowIndex, countCol + 1).setValue(currentCount + 1);
  } else {
    const newRow = invHeaders.map(h => (h === "ItemID") ? itemID : (h === "Count" ? 1 : ""));
    inventorySheet.appendRow(newRow);
  }

  writeProfile(profile, `購買商品 - ${item.ItemName}`);

  return `✅ 成功購買 ${item.ItemName}！`;
}

/**
 * 處理玩家購買並開啟十個寶箱的邏輯 (十連抽)
 * @param {string} itemID - 欲購買的寶箱 ID
 * @returns {Array<object>} - 包含 10 個獎勵物品的陣列
 */
function buyAndOpenTenItems(itemID) {
  const GUARANTEE_RARITY = 4; // 保底的星級 (4星)

  // ✅ 輔助函式：執行單次兩階段抽獎 (與 useItem 內的版本相同)
  function _performSinglePull(lootTable, allItemData, allItemHeaders) {
      const totalWeight = lootTable.reduce((sum, item) => sum + (item.Weight || 0), 0);
      if (totalWeight <= 0) throw new Error(`❌ 寶箱的總權重為 0。`);
      let random = Math.random() * totalWeight;
      let wonLoot = null;
      for (const loot of lootTable) {
        random -= (loot.Weight || 0);
        if (random <= 0) { wonLoot = loot; break; }
      }
      if (wonLoot === null) wonLoot = lootTable[lootTable.length - 1];
      const type = wonLoot.Type || "ItemRarity";
      if (type === "Currency") {
        return { type: 'currency', rewardID: wonLoot.RewardID, quantity: wonLoot.Quantity, displayName: wonLoot.DisplayName, rarity: 'currency' };
      } else {
        const selectedRarity = wonLoot.Rarity;
        const rarityCol = allItemHeaders.indexOf("Rarity");
        const potentialItems = allItemData.slice(1).filter(row => String(row[rarityCol]) === String(selectedRarity));
        if (potentialItems.length === 0) throw new Error(`❌ 在 ItemMaster 中找不到任何 Rarity 為 [${selectedRarity}] 的道具可供抽取。`);
        const wonItemRow = potentialItems[Math.floor(Math.random() * potentialItems.length)];
        return { type: 'item', itemID: wonItemRow[allItemHeaders.indexOf("ItemID")], itemName: wonItemRow[allItemHeaders.indexOf("ItemName")] || wonItemRow[allItemHeaders.indexOf("ItemID")], rarity: selectedRarity };
      }
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");
  const itemSheet = ss.getSheetByName("ItemMaster");
  const inventorySheet = ss.getSheetByName("Inventory");

  // --- 1. 讀取資料 & 檢查費用 ---
  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemRow = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === itemID);
  if (!itemRow) throw new Error("❌ 找不到此商品。");

  const item = {};
  itemHeaders.forEach((h, i) => item[h] = itemRow[i]);
  if (item.ItemType !== 'TreasureChest') throw new Error("❌ 此物品不是寶箱，無法十連抽。");

  const price = parseInt(item.BuyPrice || 0) * 10;
  const honorPrice = parseInt(item.HonorBuyPrice || 0) * 10;

  // ✅ 修正：直接讀取 Profile 工作表，確保鍵的大小寫正確，避免屬性被歸零
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  if (price > 0 && (parseInt(profile.Coins) || 0) < price) throw new Error(`⚠️ 金幣不足！十連抽需要 ${price}。`);
  if (honorPrice > 0 && (parseInt(profile.HonorPoints) || 0) < honorPrice) throw new Error(`⚠️ 榮譽點數不足！十連抽需要 ${honorPrice}。`);

  // --- 2. 執行 10 次抽獎 ---
  const lootTable = JSON.parse(item.LootTableJSON || '[]');
  const results = [];
  for (let i = 0; i < 10; i++) {
    results.push(_performSinglePull(lootTable, itemData, itemHeaders));
  }

  // --- 3. 檢查並觸發保底機制 ---
  const hasGuaranteedItem = results.some(r => r.type === 'item' && parseInt(r.rarity) >= GUARANTEE_RARITY);
  if (!hasGuaranteedItem) {
    Logger.log(`[十連抽] 未抽中 ${GUARANTEE_RARITY}★ 以上道具，觸發保底機制！`);
    const guaranteedLootTable = lootTable.filter(l => (l.Type || "ItemRarity") === "ItemRarity" && parseInt(l.Rarity) >= GUARANTEE_RARITY);
    if (guaranteedLootTable.length > 0) {
      const guaranteedItem = _performSinglePull(guaranteedLootTable, itemData, itemHeaders);
      results[9] = guaranteedItem; // 替換最後一個結果
      Logger.log(`[十連抽] 保底抽中：${guaranteedItem.rarity}★ ${guaranteedItem.itemName}`);
    }
  }

  // --- 4. 扣除費用 & 發放獎勵 ---
  if (price > 0) profile.Coins = (parseInt(profile.Coins) || 0) - price;
  if (honorPrice > 0) profile.HonorPoints = (parseInt(profile.HonorPoints) || 0) - honorPrice;

  // 處理抽中的貨幣
  results.forEach(res => {
    if (res.type === 'currency') {
      // ✅ 修正：使用從 JSON 讀取的原始 RewardID (e.g., "HonorPoints")
      profile[res.rewardID] = (parseFloat(profile[res.rewardID]) || 0) + parseFloat(res.quantity);
    }
  });

  // 批次發放道具
  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);
  const itemIDCol = invHeaders.indexOf("ItemID");
  const countCol = invHeaders.indexOf("Count");

  const inventoryMap = {};
  invRows.forEach((row, index) => {
    const currentItemID = row[itemIDCol];
    if (!inventoryMap[currentItemID]) {
      inventoryMap[currentItemID] = {
        count: parseInt(row[countCol]) || 0,
        sheetRow: index + 2
      };
    }
  });

  const itemsToAdd = {};
  results.filter(r => r.type === 'item').forEach(res => {
    itemsToAdd[res.itemID] = (itemsToAdd[res.itemID] || 0) + 1;
  });

  writeProfile(profile, `十連抽 - ${item.ItemName}`);

  const rowsToAppend = [];
  Object.entries(itemsToAdd).forEach(([id, qty]) => {
    if (inventoryMap[id] && item.IsStackable !== false) { // 假設 IsStackable
      const newCount = inventoryMap[id].count + qty;
      inventorySheet.getRange(inventoryMap[id].sheetRow, countCol + 1).setValue(newCount);
    } else {
      const newRow = invHeaders.map(h => (h === "ItemID") ? id : (h === "Count" ? qty : ""));
      rowsToAppend.push(newRow);
    }
  });

  if (rowsToAppend.length > 0) {
    inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, rowsToAppend.length, invHeaders.length).setValues(rowsToAppend);
  }

  // --- 5. 回傳結果給前端 ---
  return results.map(r => ({
    itemName: r.type === 'item' ? r.itemName : r.displayName,
    rarity: r.rarity,
    // ✅ 新增回傳資訊
    type: r.type,
    rewardID: r.rewardID,
    quantity: r.quantity
  }));
}

/**
 * 處理玩家從背包使用十個寶箱的邏輯 (十連開)
 * @param {string} itemID - 欲開啟的寶箱 ID
 * @returns {Array<object>} - 包含 10 個獎勵物品的陣列
 */
function useTenItems(itemID) {
  const GUARANTEE_RARITY = 4; // 保底的星級 (4星)

  // 輔助函式：執行單次兩階段抽獎
  function _performSinglePull(lootTable, allItemData, allItemHeaders) {
      const totalWeight = lootTable.reduce((sum, item) => sum + (item.Weight || 0), 0);
      if (totalWeight <= 0) throw new Error(`❌ 寶箱的總權重為 0。`);
      let random = Math.random() * totalWeight;
      let wonLoot = null;
      for (const loot of lootTable) {
        random -= (loot.Weight || 0);
        if (random <= 0) { wonLoot = loot; break; }
      }
      if (wonLoot === null) wonLoot = lootTable[lootTable.length - 1];
      const type = wonLoot.Type || "ItemRarity";
      if (type === "Currency") {
        return { type: 'currency', rewardID: wonLoot.RewardID, quantity: wonLoot.Quantity, displayName: wonLoot.DisplayName, rarity: 'currency' };
      } else {
        const selectedRarity = wonLoot.Rarity;
        const rarityCol = allItemHeaders.indexOf("Rarity");
        const potentialItems = allItemData.slice(1).filter(row => String(row[rarityCol]) === String(selectedRarity));
        if (potentialItems.length === 0) throw new Error(`❌ 在 ItemMaster 中找不到任何 Rarity 為 [${selectedRarity}] 的道具可供抽取。`);
        const wonItemRow = potentialItems[Math.floor(Math.random() * potentialItems.length)];
        return { type: 'item', itemID: wonItemRow[allItemHeaders.indexOf("ItemID")], itemName: wonItemRow[allItemHeaders.indexOf("ItemName")] || wonItemRow[allItemHeaders.indexOf("ItemID")], rarity: selectedRarity };
      }
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const itemSheet = ss.getSheetByName("ItemMaster");
  const profileSheet = ss.getSheetByName("Profile");
  const inventorySheet = ss.getSheetByName("Inventory");
  const logSheet = ss.getSheetByName("PlayerLog");

  // --- 1. 讀取資料 & 檢查背包數量 ---
  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);
  const itemIDColInv = invHeaders.indexOf("ItemID");
  const countColInv = invHeaders.indexOf("Count");
  
  const invIndex = invRows.findIndex(r => r[itemIDColInv] === itemID);
  if (invIndex === -1) throw new Error("❌ 背包中找不到此寶箱。");
  
  const currentCount = parseInt(invRows[invIndex][countColInv] || 0);
  if (currentCount < 10) throw new Error(`⚠️ 寶箱數量不足！十連開需要 10 個，但你只有 ${currentCount} 個。`);

  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemRow = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === itemID);
  if (!itemRow) throw new Error("❌ 找不到此商品的主資料。");

  const item = {};
  itemHeaders.forEach((h, i) => item[h] = itemRow[i]);
  if (item.ItemType !== 'TreasureChest') throw new Error("❌ 此物品不是寶箱，無法十連開。");

  // --- 2. 執行 10 次抽獎 ---
  const lootTable = JSON.parse(item.LootTableJSON || '[]');
  const results = [];
  for (let i = 0; i < 10; i++) {
    results.push(_performSinglePull(lootTable, itemData, itemHeaders));
  }

  // --- 3. 檢查並觸發保底機制 ---
  const hasGuaranteedItem = results.some(r => r.type === 'item' && parseInt(r.rarity) >= GUARANTEE_RARITY);
  if (!hasGuaranteedItem) {
    const guaranteedLootTable = lootTable.filter(l => (l.Type || "ItemRarity") === "ItemRarity" && parseInt(l.Rarity) >= GUARANTEE_RARITY);
    if (guaranteedLootTable.length > 0) {
      const guaranteedItem = _performSinglePull(guaranteedLootTable, itemData, itemHeaders);
      results[9] = guaranteedItem; // 替換最後一個結果
      Logger.log(`[十連開] 保底抽中：${guaranteedItem.rarity}★ ${guaranteedItem.itemName}`);
    }
  }

  // --- 4. 消耗背包中的寶箱 & 發放獎勵 ---
  const newCount = currentCount - 10;
  if (newCount <= 0) {
    inventorySheet.deleteRow(invIndex + 2);
  } else {
    inventorySheet.getRange(invIndex + 2, countColInv + 1).setValue(newCount);
  }

  // ✅ 修正：直接讀取 Profile 工作表，確保鍵的大小寫正確，避免屬性被歸零
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  results.filter(r => r.type === 'currency').forEach(res => {
      // ✅ 修正：使用從 JSON 讀取的原始 RewardID (e.g., "HonorPoints")
      profile[res.rewardID] = (parseFloat(profile[res.rewardID]) || 0) + parseFloat(res.quantity);
  });
  writeProfile(profile, `十連開 - ${item.ItemName}`);

  // 批次發放道具 (與 buyAndOpenTenItems 相同)
  const itemsToAdd = {};
  results.filter(r => r.type === 'item').forEach(res => { 
    itemsToAdd[res.itemID] = (itemsToAdd[res.itemID] || 0) + 1; 
  });

  const inventoryMap = {};
  const currentInvData = inventorySheet.getDataRange().getValues(); // 重新讀取，因為可能剛刪除過
  currentInvData.slice(1).forEach((row, index) => {
    inventoryMap[row[itemIDColInv]] = { count: parseInt(row[countColInv]) || 0, sheetRow: index + 2 };
  });
  const rowsToAppend = [];
  Object.entries(itemsToAdd).forEach(([id, qty]) => {
    if (inventoryMap[id] && item.IsStackable !== false) {
      inventorySheet.getRange(inventoryMap[id].sheetRow, countColInv + 1).setValue(inventoryMap[id].count + qty);
    } else {
      rowsToAppend.push(invHeaders.map(h => (h === "ItemID") ? id : (h === "Count" ? qty : "")));
    }
  });
  if (rowsToAppend.length > 0) {
    inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, rowsToAppend.length, invHeaders.length).setValues(rowsToAppend);
  }

  // --- 5. 回傳結果給前端 ---
  return results.map(r => ({
    itemName: r.type === 'item' ? r.itemName : r.displayName,
    rarity: r.rarity,
    type: r.type,
    rewardID: r.rewardID,
    quantity: r.quantity
  }));
}
/**
 * [管理頁面用] 取得所有待處理的兌換請求
 * @returns {Array<object>} - 待處理的兌換紀錄陣列
 */
function getPendingRedemptions() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RedemptionLog');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('Status');

  if (statusCol === -1) return [];

  return data
    .slice(1)
    .filter(row => row[statusCol] === 'Pending')
    .map(row => {
      const item = {};
      headers.forEach((h, i) => {
        if ((h === 'RequestDate') && row[i] instanceof Date) {
          item[h] = row[i].toISOString();
        } else {
          item[h] = row[i];
        }
      });
      return item;
    });
}

/**
 * [核銷頁面用] 根據 token 取得兌換詳情
 * @param {string} token - 核銷權杖
 * @returns {object|null} - 兌換紀錄物件
 */
function getRedemptionDetails(token) {
  if (!token) return null;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RedemptionLog');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tokenCol = headers.indexOf('Token');

  const row = data.find(r => r[tokenCol] === token);
  if (!row) return null;

  const details = {};
  headers.forEach((h, i) => {
    // 確保日期被正確序列化
    if ((h === 'RequestDate' || h === 'ProcessDate') && row[i] instanceof Date) {
      details[h] = row[i].toISOString();
    } else {
      details[h] = row[i];
    }
  });
  return details;
}

/**
 * [核銷頁面用] 處理核准或拒絕
 * @param {string} token - 核銷權杖
 * @param {string} action - 'approved' 或 'rejected'
 * @returns {string} - 結果訊息
 */
function processRedemption(token, action) {
  if (!token) throw new Error("❌ 無效的 Token。");
  if (action !== 'approved' && action !== 'rejected') throw new Error("❌ 無效的操作。");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('RedemptionLog');
  if (!sheet) throw new Error("❌ 找不到 RedemptionLog 工作表。");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tokenCol = headers.indexOf('Token');
  const statusCol = headers.indexOf('Status');
  const processDateCol = headers.indexOf('ProcessDate');

  const rowIndex = data.findIndex(r => r[tokenCol] === token);
  if (rowIndex === -1) throw new Error("❌ 找不到此兌換紀錄。");
  if (data[rowIndex][statusCol] !== 'Pending') throw new Error("⚠️ 此請求已被處理，請勿重複操作。");

  // ✅ 新增：如果拒絕，則歸還道具
  if (action === 'rejected') {
    const itemID = data[rowIndex][headers.indexOf('ItemID')];
    const itemName = data[rowIndex][headers.indexOf('ItemName')];

    if (itemID) {
      const inventorySheet = ss.getSheetByName('Inventory');
      const invData = inventorySheet.getDataRange().getValues();
      const invHeaders = invData[0];
      const invRows = invData.slice(1);
      const itemIDCol = invHeaders.indexOf('ItemID');
      const countCol = invHeaders.indexOf('Count');

      const existingItemIndex = invRows.findIndex(r => r[itemIDCol] === itemID);

      if (existingItemIndex !== -1) {
        const sheetRowIndex = existingItemIndex + 2;
        const currentCount = parseInt(invRows[existingItemIndex][countCol]) || 0;
        inventorySheet.getRange(sheetRowIndex, countCol + 1).setValue(currentCount + 1);
      } else {
        const newRow = invHeaders.map(h => (h === 'ItemID') ? itemID : (h === 'Count' ? 1 : ''));
        inventorySheet.appendRow(newRow);
      }

      // 記錄歸還事件
      const logSheet = ss.getSheetByName('PlayerLog');
      const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      const logEntry = { CreatedAt: new Date(), Type: 'system', ActionName: `兌換被拒絕 - 歸還道具`, Source: `兌換被拒絕 - 歸還 ${itemName} x1` };
      const logRow = logHeaders.map(header => logEntry[header] || '');
      logSheet.appendRow(logRow);
    }
  }

  sheet.getRange(rowIndex + 1, statusCol + 1).setValue(action === 'approved' ? 'Approved' : 'Rejected');
  sheet.getRange(rowIndex + 1, processDateCol + 1).setValue(new Date());

  if (action === 'rejected') {
    return `✅ 請求已拒絕，道具已歸還至背包。`;
  }
  return `✅ 請求已成功標示為 [${action}]！`;
}

/**
 * [核心每日自動化函式] 執行每日結算與事件檢查。
 * 應設定為每日凌晨自動執行的時間觸發器。
 * 功能：
 * 1. 檢查並觸發特殊日期事件（生日、節日）。
 * 2. 更新每日天氣並套用效果。
 * 3. 套用每日基礎消耗。
 * 4. 處理未完成任務的懲罰 (此處為範例框架，你可以將之前的懲罰邏輯整合進來)。
 */
function applyEndOfDayUpdates() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");
  const rulesSheet = ss.getSheetByName("GameRules");

  // --- 讀取玩家資料 ---
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  // --- 安全檢查：避免重複執行 ---
  const lastUpdateDate = profile.LastUpdateDate;
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  if (lastUpdateDate instanceof Date && Utilities.formatDate(lastUpdateDate, Session.getScriptTimeZone(), "yyyy/MM/dd") === todayStr) {
    Logger.log("每日結算已執行，跳過。");
    return;
  }

  // --- 1. 檢查並觸發特殊日期事件 ---
  const today = new Date();
  const todayMMDD = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd");

  // 檢查生日
  const birthday = profile.birthday;
  if (birthday instanceof Date) {
    const birthdayMMDD = Utilities.formatDate(birthday, Session.getScriptTimeZone(), "MM/dd");
    if (todayMMDD === birthdayMMDD) {
      triggerEvent('Birthday'); // 觸發事件，不用擔心重複，triggerEvent會處理
    }
  }

  // 檢查固定節日
  if (todayMMDD === "01/01") {
    triggerEvent('NewYear');
  }
  // 你可以繼續增加其他節日，例如：
  // if (todayMMDD === "12/25") {
  //   triggerEvent('Christmas');
  // }

  // --- ✅【修正】重新讀取 Profile ---
  // 因為 triggerEvent 函式會自己寫入資料，為了避免後續的每日結算使用到舊的、過時的資料（Stale Data），
  // 我們在這裡重新讀取一次玩家資料，確保接下來的計算是基於最新狀態。
  const profileRowAfterEvents = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  profileHeaders.forEach((k, i) => {
    // 更新 profile 物件，讓它反映 triggerEvent 可能造成的變更
    profile[k] = profileRowAfterEvents[i];
  });

  // 這個步驟必須在任何數值變動前執行，以捕捉角色在一天結束時的「快照」狀態
  const activeStatuses = evaluateStatusRules(profile);

  // --- 3. 處理每日基礎消耗 ---
  const rulesData = rulesSheet.getDataRange().getValues();
  rulesData.slice(1).forEach(row => {
    const [key, value] = row;
    if (!key) return;

    // 只處理每日基礎消耗
    if (key.startsWith("DailyConsumption_")) {
      const field = key.replace("DailyConsumption_", "");
      if (profile.hasOwnProperty(field)) {
        profile[field] = (parseFloat(profile[field]) || 0) + (parseFloat(value) || 0);
      }
    }
  });

  // --- 4. 套用已評估狀態的持續效果 ---
  try {
    activeStatuses.forEach(status => {
      const effects = JSON.parse(status.屬性影響JSON || '{}');
      Object.entries(effects).forEach(([field, value]) => {
        if (profile.hasOwnProperty(field)) {
          profile[field] = (parseFloat(profile[field]) || 0) + parseFloat(value);
        }
      });
    });
  } catch(e) {
    Logger.log("套用狀態效果時出錯：" + e.message);
  }

  // --- 5. 處理未完成任務懲罰 ---
  try {
    const taskSheet = ss.getSheetByName("DailyTasks");
    const taskData = taskSheet.getDataRange().getValues();
    const taskHeaders = taskData[0];
    const taskRows = taskData.slice(1);
    
    const penaltyCol = taskHeaders.indexOf("PenaltyEffects");
    const lastDoneCol = taskHeaders.indexOf("LastDoneDate");
    const taskNameCol = taskHeaders.indexOf("任務名稱");

    if (penaltyCol !== -1) {
      // 定義 "昨天" 的日期字串
      const yesterdayStr = Utilities.formatDate(new Date(new Date().getTime() - 86400000), Session.getScriptTimeZone(), "yyyy/MM/dd");

      taskRows.forEach(taskRow => {
        const penaltyJson = taskRow[penaltyCol];
        if (!penaltyJson) return; // 沒有懲罰設定，直接跳過

        const lastDoneDate = taskRow[lastDoneCol];
        const lastDoneStr = (lastDoneDate instanceof Date) 
            ? Utilities.formatDate(lastDoneDate, Session.getScriptTimeZone(), "yyyy/MM/dd") 
            : "";

        // 如果最後完成日期不是 "昨天"，就代表昨天沒有完成，應予以懲罰。
        if (lastDoneStr !== yesterdayStr) {
          try {
            const penaltyEffects = JSON.parse(penaltyJson);
            Object.entries(penaltyEffects).forEach(([field, value]) => {
              if (profile.hasOwnProperty(field)) {
                profile[field] = (parseFloat(profile[field]) || 0) + (parseFloat(value) || 0);
              }
            });
          } catch (e) {
            Logger.log(`解析任務 [${taskRow[taskNameCol]}] 的懲罰規則失敗: ${e.message}. 原始字串: "${penaltyJson}"`);
          }
        }
      });
    }
  } catch (e) {
    Logger.log("處理未完成任務懲罰失敗: " + e.message);
  }

  // --- 6. 寫回 Profile ---
  profile.LastUpdateDate = new Date(); // 記錄更新日期
  writeProfile(profile, `每日結算`);
  Logger.log(`每日結算執行完畢。`);
}

/**
 * [事件觸發器] 根據事件名稱，給予玩家對應的獎勵。
 * 此函式有防止重複觸發的機制。
 * @param {string} eventName - GameRules 中定義的事件名稱 (例如 "Birthday", "NewYear")。
 * @returns {string} 執行結果的訊息。
 */
function triggerEvent(eventName) {
  if (!eventName) return "錯誤：未提供事件名稱。";

  // --- 防止重複觸發 ---
  if (hasEventBeenTriggeredToday(eventName)) {
    const msg = `事件 '${eventName}' 今天已經觸發過了。`;
    Logger.log(msg);
    return msg;
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");
  const rulesSheet = ss.getSheetByName("GameRules");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  const rulesData = rulesSheet.getDataRange().getValues();
  const eventPrefix = `EventBonus_${eventName}_`;
  let eventApplied = false;

  rulesData.slice(1).forEach(row => {
    const [key, value] = row;
    if (key && key.startsWith(eventPrefix)) {
      const field = key.replace(eventPrefix, "");
      if (profile.hasOwnProperty(field)) {
        profile[field] = (parseFloat(profile[field]) || 0) + (parseFloat(value) || 0);
        eventApplied = true;
      }
    }
  });

  if (eventApplied) {
    // 使用 writeProfile 寫入資料，它會自動產生我們需要的日誌紀錄
    writeProfile(profile, `事件獎勵 - ${eventName}`);
    const message = `✅ 事件 [${eventName}] 觸發成功!`;
    Logger.log(message);
    return message;
  } else {
    const message = `⚠️ 在 GameRules 中找不到事件 [${eventName}] 的獎勵規則。`;
    Logger.log(message);
    return message;
  }
}

/**
 * [輔助函式] 檢查特定事件今天是否已經觸發過。
 * 透過檢查 PlayerLog 中是否有對應的 "Source" 紀錄來判斷。
 * @param {string} eventName - 要檢查的事件名稱。
 * @returns {boolean} 如果今天已觸發過則返回 true。
 */
function hasEventBeenTriggeredToday(eventName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet = ss.getSheetByName('PlayerLog');
  if (logSheet.getLastRow() < 2) return false;

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const data = logSheet.getDataRange().getValues();
  const headers = data[0];
  const sourceCol = headers.indexOf('Source');
  const dateCol = headers.indexOf('CreatedAt');
  const expectedSource = `事件獎勵 - ${eventName}`;

  // 為提高效率，只檢查最近 50 筆日誌
  const checkRange = data.slice(Math.max(1, data.length - 50));

  for (const row of checkRange) {
    const logDate = row[dateCol];
    if (logDate instanceof Date && row[sourceCol] === expectedSource) {
      if (Utilities.formatDate(logDate, Session.getScriptTimeZone(), "yyyy/MM/dd") === todayStr) {
        return true; // 找到今天觸發過的紀錄
      }
    }
  }
  return false;
}
