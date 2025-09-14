/**
 * [核心設定] 取得專案的設定值。
 * 優先從 Script Properties 讀取，若無則使用預設值。
 * 這樣可以讓每個玩家的專案副本都有自己獨立的設定，而不用修改程式碼。
 */
let _CONFIG = null; // 快取設定，避免重複讀取
function getConfig() {
  if (_CONFIG) {
    return _CONFIG;
  }

  const properties = PropertiesService.getScriptProperties();
  _CONFIG = {
    // 從指令碼屬性讀取 SPREADSHEET_ID，如果沒有，就用一個預設值 (方便開發)
    SPREADSHEET_ID: properties.getProperty('SPREADSHEET_ID') || '1OSkHqIGwq4xYEndtsrTk4Sc_EldHDeZIbvSg5L6djFs',

    // ✅ 新增：網站標題設定
    WEBSITE_TITLE: properties.getProperty('WEBSITE_TITLE') || '我的遊戲',

    // ✅【核心修改】將圖片路徑完全分開管理
    BG_IMAGE_URL: properties.getProperty('BG_IMAGE_URL') || 'https://renewmiffy.github.io/mygame/img/bg',
    CHAR_IMAGE_URL: properties.getProperty('CHAR_IMAGE_URL') || 'https://renewmiffy.github.io/mygame/img/char',
    ICON_IMAGE_URL: properties.getProperty('ICON_IMAGE_URL') || 'https://renewmiffy.github.io/mygame/img/icons'
  };

  // 在日誌中印出當前使用的設定，方便偵錯
  Logger.log(`[Config] SPREADSHEET_ID: ${_CONFIG.SPREADSHEET_ID}`);
  Logger.log(`[Config] WEBSITE_TITLE: ${_CONFIG.WEBSITE_TITLE}`);

  return _CONFIG;
}

const SPREADSHEET_ID = getConfig().SPREADSHEET_ID;

// ✅ 日期欄位一定要判斷是否為 Date 並轉為 "yyyy/MM/dd" 格式再進行比較
// 否則會導致 == 比對失敗、條件永遠不成立 顯示用途也要處理日期格式，避免出現 GMT/UTC 雜訊。
function doGet(e) { 
  const config = getConfig(); // ✅ 取得設定
  // ✅ 新增：提供一個頁面來查看所有待核銷的項目
  if (e.parameter.page === 'admin') {
    // ✅ 修正：admin.html 不需要樣板語法，直接輸出即可
    return HtmlService.createHtmlOutputFromFile('admin').setTitle(`${config.WEBSITE_TITLE} - 管理後台`).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 預設回傳遊戲主頁
  return HtmlService.createHtmlOutputFromFile('index').setTitle(config.WEBSITE_TITLE).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getInitialProfile() {
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

  // --- 2. 處理並回傳最核心的玩家資料 ---
  const profile = {};
  headers.forEach((key, i) => profile[key] = profileRow[i]);
  
  const birthdayFormatted = (profile.birthday instanceof Date)
    ? Utilities.formatDate(profile.birthday, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : (profile.birthday || '');

  const lastSurveyDateFormatted = (profile.LastSurveyDate instanceof Date)
    ? Utilities.formatDate(profile.LastSurveyDate, Session.getScriptTimeZone(), "yyyy/MM/dd")
    : (profile.LastSurveyDate || '');
  const surveyFilledToday = (lastSurveyDateFormatted === todayStr);

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
    // ✅【效能優化】只回傳檔名，由前端組合 URL
    backgroundFilename: profile.currentBackground,
    characterFilename: profile.currentCharacter,
    iconBaseUrl: getConfig().ICON_IMAGE_URL, // ✅ 新增：將圖示路徑也傳給前端
    surveyFilledToday: surveyFilledToday,
    // ✅【效能優化】將圖片基礎路徑傳給前端
    bgBaseUrl: getConfig().BG_IMAGE_URL,
    charBaseUrl: getConfig().CHAR_IMAGE_URL
  };
}

/**
 * [效能優化] 非同步獲取次要資料，例如狀態、Buff、郵件數等。
 */
function getSecondaryData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName('Profile');
  const profileRow = profileSheet.getRange(2, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profile = {};
  profileHeaders.forEach((key, i) => profile[key] = profileRow[i]);

  const activeStatuses = evaluateStatusRules(profile);
  const effectsSummary = calculateEffectsSummary(activeStatuses);
  const statusList = activeStatuses.map(s => ({ StatusName: s.狀態名稱, Effect: s.效果說明 }));
  const overrideStatus = activeStatuses.filter(s => s.CharacterOverrideFile).sort((a, b) => (a.Priority || 999) - (b.Priority || 999))[0];

  return {
    StatusList: statusList,
    effectsSummary: effectsSummary,
    characterOverrideFile: overrideStatus ? overrideStatus.CharacterOverrideFile : null,
    unreadMailCount: getUnreadMailCount()
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

  // ✅ 新增：讀取 ItemMaster 資料，用於產生更清晰的獎勵文字
  const itemSheet = ss.getSheetByName("ItemMaster");
  const itemMasterData = itemSheet.getDataRange().getValues();
  const itemMasterHeaders = itemMasterData[0];
  const itemMasterMap = {};
  itemMasterData.slice(1).forEach(row => {
    const itemID = row[itemMasterHeaders.indexOf("ItemID")];
    itemMasterMap[itemID] = { name: row[itemMasterHeaders.indexOf("ItemName")] || itemID };
  });

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

  // ✅ 建立任務完成日期的 Map，方便快速查找
  // { "TASK_ID_1": "2023/10/27", "TASK_ID_2": "2023/10/26" }
  const taskLastDoneDateMap = {};
  const doneTaskDateCol = doneTasksSheet.getRange(1, 1, 1, doneTasksSheet.getLastColumn()).getValues()[0].indexOf("LastDoneDate");
  dailyData.forEach(row => {
    const taskId = row[0];
    const lastDoneDate = row[doneTaskDateCol];
    if (taskId && lastDoneDate) taskLastDoneDateMap[taskId] = formatYMD(lastDoneDate);
  });

  // ✅ 修正：從計算「總完成數」改為計算「今天完成的任務數」
  const tasksDoneTodayCount = Object.values(taskLastDoneDateMap).filter(dateStr => dateStr === today).length;

  // ✅【核心修改】計算今天從 SkillMaster 中學習了幾個技能
  const skillLastDoneDateCol = skillHeaders.indexOf("LastDoneDate");
  const skillsLearnedTodayCount = skillData.slice(1).filter(row => {
      const lastDoneDate = row[skillLastDoneDateCol];
      return lastDoneDate && formatYMD(lastDoneDate) === today;
  }).length;


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
      // ✅ 修正：使用今天完成的任務數來判斷
      currentProgress = tasksDoneTodayCount;
      targetValue = parseInt(param);
      fulfilled = currentProgress >= targetValue;
    } else if (conditionType === "SkillsLearnedTodayCount") { // ✅【核心修改】新增條件類型
      currentProgress = skillsLearnedTodayCount;
      targetValue = parseInt(param);
      if (isNaN(targetValue)) throw new Error(`任務 [${name}] 的條件參數 "${param}" 不是一個有效的數字。`);
      fulfilled = currentProgress >= targetValue;
    } else if (conditionType === "TaskDoneToday") { // ✅ 修正：每日任務必須是今天完成的
      const lastDoneDate = taskLastDoneDateMap[param];
      const fulfilled = lastDoneDate === today;
      currentProgress = fulfilled ? 1 : 0;
      targetValue = 1;
    } else if (conditionType === "TaskDoneToday_Legacy") { // 舊的邏輯，保留以防萬一
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

    let rewardText = ""; // 用於前端顯示的格式化文字
    const rewardJsonString = getMissionField(row, "獎勵") || "{}";
    try {
      const rewardObj = JSON.parse(rewardJsonString);
      // ✅ 新增：取得欄位對照表以翻譯貨幣/屬性名稱
      const fieldMap = getCachedFieldMap();

      // 檢查確保解析出來的是一個物件
      if (typeof rewardObj === 'object' && rewardObj !== null && !Array.isArray(rewardObj)) {
        // ✅ 修正：正確處理道具獎勵的顯示
        rewardText = Object.entries(rewardObj).map(([key, val]) => {
          if (key === 'Items' && Array.isArray(val)) {
            return val.map(item => {
              const itemName = itemMasterMap[item.ItemID]?.name || item.ItemID;
              return `${itemName} x${item.Quantity}`;
            }).join(', ');
          }
          // ✅ 修正：對於非 Items 的獎勵，使用對照表翻譯
          return `+${val} ${mapFieldToName(key, fieldMap)}`;
        }).join(" / ");
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
      rewardJson: rewardJsonString, // ✅ 新增：回傳原始 JSON 字串供編輯器使用
      rewardText: rewardText,
      fulfilled: fulfilled,
      claimed: claimed,
      displayOrder: displayOrder,
      rowIndex: i,
      conditionType: conditionType, // ✅ 新增：回傳條件類型
      conditionParam: param,        // ✅ 新增：回傳條件參數
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
    // ✅ 使用新的獎勵處理函式
    const { message } = applyRewards(profile, rewardObj, `任務 - ${getMissionField(mission, "任務名稱")}`);
    rewardText = message;

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

    // --- ✅【效能優化】一次性讀取所有需要的資料 ---
    const taskData = taskSheet.getDataRange().getValues();
    const taskHeaders = taskData[0];
    const taskRows = taskData.slice(1);

    const profileData = profileSheet.getDataRange().getValues();
    const profileHeaders = profileData[0];
    const profileRow = profileData[1];
    const profile = {};
    profileHeaders.forEach((key, i) => profile[key] = profileRow[i]);

    // --- 執行任務邏輯 ---
    const taskIndex = taskRows.findIndex(r => r[0] === taskId);
    if (taskIndex === -1) throw new Error("❌ 找不到此任務");

    const taskRowToUpdate = taskRows[taskIndex];
    const lastDoneDateCol = taskHeaders.indexOf("LastDoneDate");
    const streakCol = taskHeaders.indexOf("StreakCount");
    const totalCol = taskHeaders.indexOf("TotalDoneCount");
    const effectsCol = taskHeaders.indexOf("Effects");
    const nameCol = taskHeaders.indexOf("任務名稱");

    const lastDate = taskRowToUpdate[lastDoneDateCol];
    const todayStr = formatYMD(new Date());
    if (lastDate instanceof Date && formatYMD(lastDate) === todayStr) {
        throw new Error("⚠️ 此任務今日已完成");
    }

    const effects = JSON.parse(taskRowToUpdate[effectsCol] || "{}");
    const taskName = taskRowToUpdate[nameCol] || "未知任務";
    const rewardResult = applyRewards(profile, effects, "日常任務 - " + taskName);

    const lastDateStr = (lastDate instanceof Date) ? formatYMD(lastDate) : "";
    const yesterdayStr = formatYMD(new Date(new Date().getTime() - 86400000));
    const currentStreak = parseInt(taskRowToUpdate[streakCol] || 0);

    taskRowToUpdate[lastDoneDateCol] = new Date();
    taskRowToUpdate[streakCol] = (lastDateStr === yesterdayStr) ? currentStreak + 1 : 1;
    taskRowToUpdate[totalCol] = (parseInt(taskRowToUpdate[totalCol] || 0)) + 1;

    // --- ✅【效能優化】一次性寫回更新後的任務資料 ---
    taskSheet.getRange(taskIndex + 2, 1, 1, taskHeaders.length).setValues([taskRowToUpdate]);

    return {
        message: `✅ ${taskName} 完成！`,
        penaltyInfo: rewardResult.penaltyInfo
    };
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

    // --- ✅【效能優化】一次性讀取所有需要的資料 ---
    const skillData = skillSheet.getDataRange().getValues();
    const skillHeaders = skillData[0];
    const skillRows = skillData.slice(1);

    const profileData = profileSheet.getDataRange().getValues();
    const profileHeaders = profileData[0];
    const profileRow = profileData[1];
    const profile = {};
    profileHeaders.forEach((key, i) => profile[key] = profileRow[i]);

    // --- 執行學習邏輯 ---
    const skillIndex = skillRows.findIndex(r => r[0] === skillId);
    if (skillIndex === -1) throw new Error("❌ 找不到此技能");

    const skillRowToUpdate = skillRows[skillIndex];
    const lastDoneDateCol = skillHeaders.indexOf("LastDoneDate");
    const streakCol = skillHeaders.indexOf("StreakCount");
    const totalCol = skillHeaders.indexOf("TotalDoneCount");

    const lastDate = skillRowToUpdate[lastDoneDateCol];
    const todayStr = formatYMD(new Date());
    if (lastDate instanceof Date && formatYMD(lastDate) === todayStr) throw new Error("⚠️ 今天已學過此技能");

    const effects = JSON.parse(skillRowToUpdate[skillHeaders.indexOf("Effects")] || "{}");
    const skillName = skillRowToUpdate[skillHeaders.indexOf("SkillName")] || "未知技能";
    const rewardResult = applyRewards(profile, effects, "學習 - " + skillName);

    const lastDateStr = (lastDate instanceof Date) ? formatYMD(lastDate) : "";
    const yesterdayStr = formatYMD(new Date(new Date().getTime() - 86400000));

    skillRowToUpdate[lastDoneDateCol] = new Date();
    skillRowToUpdate[streakCol] = (lastDateStr === yesterdayStr) ? (parseInt(skillRowToUpdate[streakCol] || 0)) + 1 : 1;
    skillRowToUpdate[totalCol] = (parseInt(skillRowToUpdate[totalCol] || 0)) + 1;

    // --- ✅【效能優化】一次性寫回更新後的技能資料 ---
    skillSheet.getRange(skillIndex + 2, 1, 1, skillHeaders.length).setValues([skillRowToUpdate]);

    return {
        message: `🎓 學習 ${skillName} 完成`,
        penaltyInfo: rewardResult.penaltyInfo
    };
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
        CharacterOverrideFile: rule["CharacterOverrideFile"], // ✅ 新增
        // ✅ 修正：讀取所有獎勵加成相關欄位
        CoinBonusPercent: rule["CoinBonusPercent"],
        HonorBonusPercent: rule["HonorBonusPercent"],
        ShopDiscountPercent: rule["ShopDiscountPercent"],
        GlobalRewardModifier: rule["GlobalRewardModifier"]
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

  // ✅ 新增：讀取並回傳當前生效的 Buff
  const activeBuffs = getActiveBuffs();
  const buffList = activeBuffs.map(b => ({ BuffName: b.BuffName, ExpiryDate: b.ExpiryDate.toISOString() }));

  // ✅ 新增：計算並回傳加成效果總結
  const effectsSummary = calculateEffectsSummary(activeStatuses);
  Logger.log(`[getQuickStatus] 計算出的效果總結: ${JSON.stringify(effectsSummary)}`);

  // ✅ 新增：狀態圖片更換邏輯，與 getProfileData() 同步
  const overrideStatus = activeStatuses
    .filter(s => s.CharacterOverrideFile)
    .sort((a, b) => (a.Priority || 999) - (b.Priority || 999))[0];

  let finalCharacterFile = profile.currentCharacter;
  if (overrideStatus) {
    finalCharacterFile = overrideStatus.CharacterOverrideFile;
  }

  return {
    // ✅ 修正：確保回傳所有 profile 屬性，與 getProfileData() 一致
    playerName: profile.PlayerName,
    birthday: (profile.birthday instanceof Date) ? Utilities.formatDate(profile.birthday, Session.getScriptTimeZone(), "yyyy/MM/dd") : (profile.birthday || ''),
    coins: profile.Coins || 0,
    honorPoints: profile.HonorPoints || 0,
    cleanliness: profile.Cleanliness || 0,
    mood: profile.Mood || 0,
    energy: profile.Energy || 0,
    health: profile.Health || 0,
    selfDiscipline: profile.SelfDiscipline || 0,
    StatusList: statusList,
    // ✅ 修正：使用設定檔中的 URL
    backgroundUrl: `${getConfig().BG_IMAGE_URL}/${profile.currentBackground}`,
    characterUrl: `${getConfig().CHAR_IMAGE_URL}/${finalCharacterFile}`,
    effectsSummary: effectsSummary,
    unreadMailCount: getUnreadMailCount() // ✅ 新增：刷新時也回傳未讀郵件數
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");
  const inventorySheet = ss.getSheetByName("Inventory");
  const itemSheet = ss.getSheetByName("ItemMaster");
  const logSheet = ss.getSheetByName("PlayerLog");

  // ✅【關鍵修正】將這段被誤刪的程式碼加回來，以正確讀取玩家資料
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  // --- ✅【效能優化】一次性讀取所有需要的資料 ---
  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemRow = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === itemID);
  if (!itemRow) throw new Error("❌ 找不到道具主資料 (ItemMaster)");

  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);
  const invIndex = invRows.findIndex(row => row[invHeaders.indexOf("ItemID")] === itemID);
  if (invIndex === -1) throw new Error("❌ 背包中找不到此道具");

  quantity = parseInt(quantity) || 1;
  if (quantity <= 0) throw new Error("⚠️ 使用數量必須大於 0");
  const count = parseInt(invRows[invIndex][invHeaders.indexOf("Count")]);
  if (count < quantity) throw new Error(`⚠️ 數量不足，需要 ${quantity} 個，但你只有 ${count} 個。`);
  
  const itemMaster = {};
  itemHeaders.forEach((k, i) => itemMaster[k] = itemRow[i]);
  const itemName = itemMaster.ItemName || itemID;
  const itemType = itemMaster.ItemType || 'Consumable'; // 預設為消耗品
  
  // ----------------------------------------------------------------
  // ✅ 新增：處理固定內容禮包 (FixedBundle)
  // ----------------------------------------------------------------
  if (itemType === 'FixedBundle') {
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
    const itemIDCol = invHeaders.indexOf("ItemID");
    const countCol = invHeaders.indexOf("Count");
    const inventoryMap = {};
    invRows.forEach((row, index) => { inventoryMap[row[itemIDCol]] = { count: parseInt(row[countCol]) || 0, sheetRow: index + 2 }; });
    
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
    else { inventorySheet.getRange(invIndex + 2, invHeaders.indexOf("Count") + 1).setValue(count - quantity); }
    
    writeProfile(profile, `開啟禮包 - ${itemName}`);
    // ✅ 修正：回傳結構化資料
    return {
      message: `✅ 已開啟 ${itemName} x${quantity}！獲得：${itemsGranted.join(', ')}`,
      profileData: getQuickStatus() // 順便回傳最新狀態
    };
  }
  // ----------------------------------------------------------------
  // ✅ 新增：處理時效性 Buff 道具 (BuffItem)
  // ----------------------------------------------------------------
  else if (itemType === 'BuffItem') {
    const effectJson = itemMaster.Effect || '{}';
    let buffInfo;
    try {
      buffInfo = JSON.parse(effectJson);
      if (buffInfo.type !== 'buff' || !buffInfo.duration_hours || !buffInfo.effects) {
        throw new Error("Effect JSON 格式不符，缺少 type, duration_hours 或 effects。");
      }
    } catch (e) {
      throw new Error(`❌ 無法解析 Buff 道具 [${itemName}] 的效果設定 (Effect): ${e.message}`);
    }
    
    // ✅ 修正：如果工作表不存在，就自動建立它
    let buffSheet = ss.getSheetByName('ActiveBuffs');
    if (!buffSheet) {
      buffSheet = ss.insertSheet('ActiveBuffs');
      buffSheet.appendRow(['BuffID', 'BuffName', 'EffectJSON', 'ExpiryDate', 'AppliedDate']);
    }
    
    const now = new Date();
    const expiryDate = new Date(now.getTime() + buffInfo.duration_hours * 60 * 60 * 1000);
    
    buffSheet.appendRow([Utilities.getUuid(), itemName, JSON.stringify(buffInfo.effects), expiryDate, now]);
    
    // 消耗道具
    if (count - quantity <= 0) { inventorySheet.deleteRow(invIndex + 2); }
    else { inventorySheet.getRange(invIndex + 2, invHeaders.indexOf("Count") + 1).setValue(count - quantity); }

    // ✅ 修正：回傳結構化資料
    return {
      message: `✅ 已使用 ${itemName}！效果將持續 ${buffInfo.duration_hours} 小時。`,
      profileData: getQuickStatus()
    };
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
        const rewardsToApply = { [field]: parseFloat(value) * quantity };
        // ✅ 使用新的獎勵處理函式
        const { finalRewards, message } = applyRewards(profile, rewardsToApply, `開啟貨幣包 - ${itemName}`);
        effectApplied = true;
        rewardsMessage.push(message);
      });
    } catch (e) {
      throw new Error(`❌ 無法解析貨幣包 [${itemName}] 的效果設定 (Effect): ${e.message}`);
    }
    
    if (!effectApplied) { throw new Error(`❌ 貨幣包 [${itemName}] 的效果設定無效。`); }
    
    if (count - quantity <= 0) { inventorySheet.deleteRow(invIndex + 2); }
    else { inventorySheet.getRange(invIndex + 2, invHeaders.indexOf("Count") + 1).setValue(count - quantity); }
    
    // ✅ 修正：回傳結構化資料
    return {
      message: `✅ 已開啟 ${itemName} x${quantity}！獲得：${rewardsMessage.join(', ')}`,
      profileData: getQuickStatus()
    };
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
    const subject = `[遊戲獎勵兌換] ${profile.PlayerName} 請求兌換：${itemName}`;
    const body = `
      <h3>您好！</h3>
      <p>玩家 <strong>${profile.PlayerName}</strong> 在遊戲中請求兌換以下實體獎勵：</p>
      <p style="font-size: 18px; font-weight: bold;">獎勵名稱： ${itemName}</p>
      <p>請點擊以下連結，前往管理後台查看所有待處理的項目：</p>
      <p><a href="${adminUrl}" style="font-size: 16px; padding: 10px 15px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px;">前往管理後台</a></p>
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
        // ✅ 新增：處理 "All" 這個特殊關鍵字
        if (field === 'All') {
          const attributesToBoost = ['Cleanliness', 'Mood', 'Energy', 'Health', 'SelfDiscipline'];
          const boostValue = parseFloat(value) * quantity;
          attributesToBoost.forEach(attr => {
            if (profile.hasOwnProperty(attr)) {
              profile[attr] = (parseFloat(profile[attr]) || 0) + boostValue;
            }
          });
        } else if (profile.hasOwnProperty(field)) {
          // 原本的邏輯：處理單一屬性
          profile[field] = (parseFloat(profile[field]) || 0) + (parseFloat(value) * quantity);
        }
      });
    } catch (e) { /* 忽略錯誤 */ }
  }

  if (count - quantity <= 0) {
    inventorySheet.deleteRow(invIndex + 2);
  } else {
    inventorySheet.getRange(invIndex + 2, invHeaders.indexOf("Count") + 1).setValue(count - quantity);
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
    // ✅ 修正：回傳結構化資料
    return {
      message: "✅ 兌換請求已發送！請等待家長為您核准。",
      profileData: getQuickStatus()
    };
  } else {
    writeProfile(profile, `使用道具 - ${itemName}`);
    // ✅ 修正：回傳結構化資料
    return {
      message: `✅ 已使用 ${itemName} x${quantity}！`,
      profileData: getQuickStatus()
    };
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
        HonorBuyPrice: parseInt(item.HonorBuyPrice || 0),
        Rarity: item.Rarity || 0 // ✅ 新增：回傳稀有度供前端排序
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

  // --- ✅【效能優化】一次性讀取所有需要的資料 ---
  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemRow = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === itemID);
  if (!itemRow) throw new Error("❌ 找不到此商品。");

  // --- 處理商品資料 ---
  const item = {};
  itemHeaders.forEach((h, i) => item[h] = itemRow[i]);
  if (item.IsPurchasable !== true) throw new Error("❌ 此商品不可購買。");

  const price = parseInt(item.BuyPrice || 0);
  const honorPrice = parseInt(item.HonorBuyPrice || 0);

  // --- 讀取 Profile 並計算費用 ---
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  const activeStatuses = evaluateStatusRules(profile);
  const effectsSummary = calculateEffectsSummary(activeStatuses);
  const discountRate = (100 - (effectsSummary.shopDiscount || 0)) / 100;

  const finalPrice = Math.round(price * discountRate);
  const finalHonorPrice = Math.round(honorPrice * discountRate);

  if (finalPrice > 0 && (parseInt(profile.Coins) || 0) < finalPrice) {
    throw new Error(`⚠️ 金幣不足！需要 ${finalPrice}。`);
  }
  if (finalHonorPrice > 0 && (parseInt(profile.HonorPoints) || 0) < finalHonorPrice) {
    throw new Error(`⚠️ 榮譽點數不足！需要 ${finalHonorPrice}。`);
  }

  // --- 扣錢 & 發放道具 ---
  if (price > 0) profile.Coins = (parseInt(profile.Coins) || 0) - finalPrice;
  if (honorPrice > 0) profile.HonorPoints = (parseInt(profile.HonorPoints) || 0) - finalHonorPrice;

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
 * 處理玩家從背包使用十個寶箱的邏輯 (十連開)
 * @param {string} itemID - 欲開啟的寶箱 ID
 * @returns {Array<object>} - 包含 10 個獎勵物品的陣列
 */
function useTenItems(itemID) {
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
        if ((h === 'RequestDate' || h === 'ProcessDate') && row[i] instanceof Date) {
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
  // ✅ 新增：讀取日常和學習工作表
  const taskSheet = ss.getSheetByName("DailyTasks");
  const skillSheet = ss.getSheetByName("SkillMaster");

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

  // --- ✅【新功能】處理昨日無活動懲罰 ---
  try {
    const yesterdayStr = formatYMD(new Date(new Date().getTime() - 86400000));
    let wasActiveYesterday = false;

    // ✅【核心修改】只檢查學習紀錄
    if (skillSheet && skillSheet.getLastRow() > 1) {
      const skillData = skillSheet.getDataRange().getValues();
      const lastDoneCol = skillData[0].indexOf("LastDoneDate");
      if (skillData.slice(1).some(row => formatYMD(row[lastDoneCol]) === yesterdayStr)) {
        wasActiveYesterday = true;
      }
    }

    // 如果昨天完全沒有活動，則套用懲罰
    if (!wasActiveYesterday) {
      const inactivityPenalty = parseFloat(rulesData.find(row => row[0] === 'Penalty_Inactivity_SelfDiscipline')?.[1] || 0);
      if (inactivityPenalty < 0) {
        profile.SelfDiscipline = (parseFloat(profile.SelfDiscipline) || 0) + inactivityPenalty;
        Logger.log(`[每日結算] 昨日無任何學習活動，自律 ${inactivityPenalty}`);
      }
    }
  } catch (e) { Logger.log(`處理昨日無活動懲罰時出錯: ${e.message}`); }

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

/**
 * [輔助函式] 根據欄位對照表，將英文欄位名轉為中文。
 * @param {string} field - 英文欄位名。
 * @param {object} fieldMap - 從 getFieldMapping() 取得的對照表。
 * @returns {string} - 中文名稱或原始英文名稱。
 */
function mapFieldToName(field, fieldMap) {
  return fieldMap[field] || field;
}

let cachedFieldMap = null; // ✅ 新增：快取欄位對照表，避免重複讀取
function getCachedFieldMap() {
  if (!cachedFieldMap) cachedFieldMap = getFieldMapping();
  return cachedFieldMap;
}

/**
 * [新輔助函式] 將指定道具發放到玩家背包。
 * @param {object} itemsToAdd - 一個物件，鍵為 ItemID，值為數量。例如：{"POTION_01": 2, "SCROLL_01": 1}
 * @returns {string} - 描述獲得物品的訊息，例如 "藥水 x2, 卷軸 x1"。
 */
function grantItemsToInventory(itemsToAdd) {
  if (!itemsToAdd || Object.keys(itemsToAdd).length === 0) {
    return "";
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const inventorySheet = ss.getSheetByName("Inventory");
  const itemSheet = ss.getSheetByName("ItemMaster");

  // 讀取道具主資料以解析道具名稱
  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemMasterMap = {};
  itemData.slice(1).forEach(row => {
    const itemID = row[itemHeaders.indexOf("ItemID")];
    itemMasterMap[itemID] = { name: row[itemHeaders.indexOf("ItemName")] || itemID };
  });

  // 批次發放道具的邏輯
  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];
  const invRows = invData.slice(1);
  const itemIDCol = invHeaders.indexOf("ItemID");
  const countCol = invHeaders.indexOf("Count");

  const inventoryMap = {};
  invRows.forEach((row, index) => {
    if (!inventoryMap[row[itemIDCol]]) { // 只記錄第一個堆疊
      inventoryMap[row[itemIDCol]] = { count: parseInt(row[countCol]) || 0, sheetRow: index + 2 };
    }
  });

  const rowsToAppend = [];
  Object.entries(itemsToAdd).forEach(([id, qty]) => {
    if (inventoryMap[id]) { // 假設所有獎勵道具都可堆疊
      inventorySheet.getRange(inventoryMap[id].sheetRow, countCol + 1).setValue(inventoryMap[id].count + qty);
    } else {
      rowsToAppend.push(invHeaders.map(h => (h === "ItemID") ? id : (h === "Count" ? qty : "")));
    }
  });

  if (rowsToAppend.length > 0) {
    inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, rowsToAppend.length, invHeaders.length).setValues(rowsToAppend);
  }

  // 產生回傳訊息
  return Object.entries(itemsToAdd).map(([id, qty]) => `${itemMasterMap[id]?.name || id} x${qty}`).join(', ');
}

/**
 * [核心獎勵處理函式] 根據玩家狀態，計算並套用最終獎勵。
 * @param {object} profile - 玩家的 profile 物件 (會被直接修改)。
 * @param {object} rewards - 原始獎勵物件，例如 { "Coins": 100, "Health": 5 }。
 * @param {string} source - 獎勵來源的文字描述，用於日誌記錄。
 * @returns {{finalRewards: object, message: string}} - 包含最終獎勵值和描述訊息的物件。
 */
function applyRewards(profile, rewards, source) {
  // ✅ 新增：準備回傳給前端的懲罰資訊
  let penaltyInfo = null;
  const activeStatuses = evaluateStatusRules(profile);
  const effectsSummary = calculateEffectsSummary(activeStatuses);

  if (effectsSummary.globalRewardModifier < 1) {
    const penaltyReasonStatus = activeStatuses.find(s => s.GlobalRewardModifier && parseFloat(s.GlobalRewardModifier) < 1);
    penaltyInfo = {
      wasApplied: true,
      reason: penaltyReasonStatus ? penaltyReasonStatus.狀態名稱 : "狀態不佳"
    };
  }

  const finalRewards = {};
  const messageParts = [];
  const fieldMap = getCachedFieldMap(); // ✅ 取得欄位對照表

  // --- 1. 處理道具獎勵 ---
  if (rewards.Items && Array.isArray(rewards.Items)) {
    const itemsToAdd = {};
    rewards.Items.forEach(item => {
      if (item.ItemID && item.Quantity > 0) {
        itemsToAdd[item.ItemID] = (itemsToAdd[item.ItemID] || 0) + parseInt(item.Quantity);
      }
    });
    if (Object.keys(itemsToAdd).length > 0) {
      const grantedItemsMessage = grantItemsToInventory(itemsToAdd);
      if (grantedItemsMessage) messageParts.push(grantedItemsMessage);
    }
  }

  // --- 2. 處理數值獎勵 (金幣、屬性等) ---
  Object.entries(rewards).forEach(([field, value]) => {
    if (field === 'Items') return; // 忽略道具陣列，前面已處理
    let finalValue = parseFloat(value) || 0;

    // 只對非道具的獎勵套用加成
    if (profile.hasOwnProperty(field)) {
      // 1. 套用全域修正 (例如 "髒兮兮" 懲罰)
      finalValue *= effectsSummary.globalRewardModifier;

      // 2. 套用特定類型的百分比加成
      if (field === 'Coins') {
        finalValue *= (1 + (effectsSummary.coinBonus / 100));
      } else if (field === 'HonorPoints') {
        finalValue *= (1 + (effectsSummary.honorBonus / 100));
      }

      // ✅【核心修正】只對金幣進行四捨五入，其他屬性保留小數點
      if (field === 'Coins') {
        finalValue = Math.round(finalValue);
      } else {
        // 保留到小數點後兩位
        finalValue = Math.round(finalValue * 100) / 100;
      }

      // 更新 profile
      profile[field] = (parseFloat(profile[field]) || 0) + finalValue;
      finalRewards[field] = finalValue;
      messageParts.push(`${mapFieldToName(field, fieldMap) || field} +${finalValue}`); // ✅ 修正呼叫方式
    }
  });

  // 使用 writeProfile 統一寫入，它會自動處理日誌
  if (Object.keys(finalRewards).length > 0) {
    writeProfile(profile, source);
  }

  return {
    finalRewards: finalRewards,
    message: messageParts.join(', '),
    penaltyInfo: penaltyInfo // ✅ 新增
  };
}

/**
 * [輔助函式] 取得所有未過期的 Buff，並順便清理已過期的。
 * @returns {Array<object>} - 未過期的 Buff 物件陣列。
 */
function getActiveBuffs() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const buffSheet = ss.getSheetByName('ActiveBuffs');
  if (!buffSheet || buffSheet.getLastRow() < 2) {
    return [];
  }

  const data = buffSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const now = new Date();
  const activeBuffs = [];
  const expiredRowIndices = [];

  const expiryCol = headers.indexOf('ExpiryDate');

  rows.forEach((row, index) => {
    const expiryDate = new Date(row[expiryCol]);
    if (expiryDate > now) {
      const buff = {};
      headers.forEach((h, i) => buff[h] = row[i]);
      buff.ExpiryDate = expiryDate; // 確保是 Date 物件
      activeBuffs.push(buff);
    } else {
      expiredRowIndices.push(index + 2); // 記錄試算表中的實際列號
    }
  });

  // 從後往前刪除，避免索引錯亂
  expiredRowIndices.reverse().forEach(rowIndex => {
    buffSheet.deleteRow(rowIndex);
  });

  return activeBuffs;
}

/**
 * [核心效果計算函式] 根據生效中的狀態，匯總所有加成/折扣效果。
 * @param {Array<object>} activeStatuses - 從 evaluateStatusRules() 得到的生效狀態陣列。
 * @returns {object} - 一個包含所有效果總和的物件。
 *                   例如：{ coinBonus: 15, honorBonus: 0, shopDiscount: 5, globalRewardModifier: 0.8 }
 */
function calculateEffectsSummary(activeStatuses) {
  const summary = {
    coinBonus: 0,
    honorBonus: 0,
    shopDiscount: 0,
    globalRewardModifier: 1.0,
  };

  if (!activeStatuses || activeStatuses.length === 0) {
    return summary;
  }

  activeStatuses.forEach(status => {
    // 屬性影響JSON是每日結算用的，這裡不處理。我們處理新欄位。
    summary.coinBonus += parseFloat(status.CoinBonusPercent) || 0;
    summary.honorBonus += parseFloat(status.HonorBonusPercent) || 0; // 來自角色狀態的加成
    summary.shopDiscount += parseFloat(status.ShopDiscountPercent) || 0; // 來自角色狀態的折扣

    const modifier = parseFloat(status.GlobalRewardModifier);
    if (!isNaN(modifier) && modifier > 0) {
      summary.globalRewardModifier *= modifier;
    }
  });

  // ✅ 新增：讀取並加入 ActiveBuffs 的效果
  const activeBuffs = getActiveBuffs();
  activeBuffs.forEach(buff => {
    try {
      const effects = JSON.parse(buff.EffectJSON || '{}');
      Object.entries(effects).forEach(([key, value]) => {
        // ✅ 修正：直接根據 key 將效果累加到 summary 物件中
        if (key.endsWith('Percent')) {
          if (key === 'ShopDiscountPercent') summary.shopDiscount += parseFloat(value) || 0;
          if (key === 'CoinBonusPercent') summary.coinBonus += parseFloat(value) || 0;
          if (key === 'HonorBonusPercent') summary.honorBonus += parseFloat(value) || 0;
        }
      });
    } catch (e) {
      Logger.log(`解析 Buff [${buff.BuffName}] 的效果時出錯: ${e.message}`);
    }
  });

  return summary;
}

/**
 * [郵件系統] 獲取玩家的郵件列表 (未過期、未領取)。
 * @returns {Array<object>} 郵件物件陣列。
 */
function getMailList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const mailSheet = ss.getSheetByName('Mailbox');
  if (!mailSheet || mailSheet.getLastRow() < 2) return [];

  const data = mailSheet.getDataRange().getValues();
  const headers = data[0];
  const now = new Date();

  // 取得道具名稱對照，用於產生獎勵文字
  const itemSheet = ss.getSheetByName("ItemMaster");
  const itemMasterData = itemSheet.getDataRange().getValues();
  const itemMasterHeaders = itemMasterData[0];
  const itemMasterMap = {};
  itemMasterData.slice(1).forEach(row => {
    const itemID = row[itemMasterHeaders.indexOf("ItemID")];
    itemMasterMap[itemID] = { name: row[itemMasterHeaders.indexOf("ItemName")] || itemID };
  });

  const mails = data.slice(1).map((row, index) => {
    const mail = {};
    headers.forEach((h, i) => mail[h] = row[i]);
    mail.sheetRow = index + 2; // 記下實際列號方便更新
    return mail;
  }).filter(mail => {
    const isClaimed = mail.IsClaimed === true;
    const expiryDate = mail.ExpiryDate instanceof Date ? mail.ExpiryDate : null;
    const isExpired = expiryDate ? expiryDate < now : false;
    return !isClaimed && !isExpired;
  }).map(mail => {
    let rewardText = "";
    try {
      const rewardObj = JSON.parse(mail.RewardJSON || '{}');
      if (Object.keys(rewardObj).length > 0) { // ✅ 關鍵修正：只有在獎勵物件不是空的時候才產生文字
        rewardText = Object.entries(rewardObj).map(([key, val]) => {
          if (key === 'Items' && Array.isArray(val)) {
            return val.map(item => `${itemMasterMap[item.ItemID]?.name || item.ItemID} x${item.Quantity}`).join(', ');
          }
          return `${mapFieldToName(key, getCachedFieldMap())} +${val}`;
        }).join(" / ");
      }
    } catch (e) { /* 忽略解析錯誤 */ }

    return {
      id: mail.MailID,
      title: mail.Title,
      message: mail.Message,
      sentDate: mail.SentDate instanceof Date ? Utilities.formatDate(mail.SentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd') : '',
      isRead: mail.IsRead === true,
      hasReward: rewardText !== "", // ✅ 關鍵修正：判斷獎勵文字是否為空
      rewardText: rewardText
    };
  }).sort((a, b) => new Date(b.sentDate) - new Date(a.sentDate)); // 按發送日期降序

  return mails;
}

/**
 * [郵件系統] 玩家領取郵件獎勵。
 * @param {string} mailID - 要領取的郵件 ID。
 * @returns {object} 包含訊息和最新玩家狀態的物件。
 */
function claimMailReward(mailID) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 等待最多 10 秒

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const mailSheet = ss.getSheetByName('Mailbox');
    const data = mailSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('MailID');
    const claimedCol = headers.indexOf('IsClaimed');

    const mailIndex = data.slice(1).findIndex(row => row[idCol] === mailID);
    if (mailIndex === -1) throw new Error("❌ 找不到該郵件。");

    const mailRow = data[mailIndex + 1];
    if (mailRow[claimedCol] === true) throw new Error("⚠️ 您已經領取過此獎勵。");

    const profileSheet = ss.getSheetByName("Profile");
    const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
    const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
    const profile = {};
    profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

    const rewardJSON = mailRow[headers.indexOf('RewardJSON')];
    const rewardObj = JSON.parse(rewardJSON || '{}');

    const { message } = applyRewards(profile, rewardObj, `郵件獎勵 - ${mailRow[headers.indexOf('Title')]}`);

    // 更新郵件狀態
    mailSheet.getRange(mailIndex + 2, claimedCol + 1).setValue(true);
    mailSheet.getRange(mailIndex + 2, headers.indexOf('IsRead') + 1).setValue(true);

    return { message: `✅ 獎勵已領取：${message}`, profileData: getQuickStatus() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * [郵件系統] 將郵件標示為已讀，並回傳最新的未讀數量。
 * @param {string} mailID - 郵件 ID。
 * @returns {number} 最新的未讀郵件數量。
 */
function markMailAsReadAndGetCount(mailID) {
  Logger.log(`--- [markMailAsReadAndGetCount] 開始執行，目標 MailID: "${mailID}" (類型: ${typeof mailID}) ---`);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const mailSheet = ss.getSheetByName('Mailbox');
    const data = mailSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('MailID');
    const readCol = headers.indexOf('IsRead');

    if (idCol === -1 || readCol === -1) {
      Logger.log(`❌ 錯誤：在 Mailbox 工作表中找不到 'MailID' 或 'IsRead' 欄位。`);
      return getUnreadMailCount();
    }

    const mailIndex = data.slice(1).findIndex(row => {
      const sheetMailID = String(row[idCol] || '').trim();
      const frontendMailID = String(mailID || '').trim();
      return sheetMailID === frontendMailID;
    });

    if (mailIndex !== -1) {
      Logger.log(`✅ 成功找到匹配的郵件，位於資料陣列索引 ${mailIndex} (工作表第 ${mailIndex + 2} 列)。正在將 IsRead 設為 true...`);
      mailSheet.getRange(mailIndex + 2, readCol + 1).setValue(true);
      SpreadsheetApp.flush(); // 強制寫入
      Utilities.sleep(1000);  // 等待同步
    } else {
      Logger.log(`❌ 警告：在 Mailbox 工作表中找不到 MailID 為 "${mailID}" 的郵件。`);
    }
    return getUnreadMailCount();
  } catch (e) {
    Logger.log(`❌ 在 markMailAsReadAndGetCount 中發生嚴重錯誤: ${e.message}`);
    return getUnreadMailCount(); // 即使出錯，也回傳當前計數，避免前端崩潰
  }
}

/**
 * [郵件系統] 將郵件標示為已讀。
 * @param {string} mailID - 郵件 ID。
 */
function markMailAsRead(mailID) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const mailSheet = ss.getSheetByName('Mailbox');
  const data = mailSheet.getDataRange().getValues();
  const idCol = data[0].indexOf('MailID');
  const readCol = data[0].indexOf('IsRead');

  const mailIndex = data.slice(1).findIndex(row => row[idCol] === mailID);
  if (mailIndex !== -1) {
    mailSheet.getRange(mailIndex + 2, readCol + 1).setValue(true);
  }
}

/**
 * [郵件系統] 獲取未讀郵件的數量。
 */
function getUnreadMailCount() {
  const mails = getMailList();
  return mails.filter(m => !m.isRead).length;
}

/**
 * [管理後台用] 發送一封系統郵件給玩家。
 * @param {string} title - 郵件標題。
 * @param {string} message - 郵件內文。
 * @param {string} rewardJson - 獎勵內容的 JSON 字串。
 * @param {number} expiryDays - 郵件的有效天數。
 * @returns {string} 執行結果訊息。
 */
function sendAdminMail(title, message, rewardJson, expiryDays) {
  // 基本驗證
  if (!title || !message) {
    throw new Error("❌ 標題和內文為必填項目。");
  }

  let parsedReward = {};
  if (rewardJson && rewardJson.trim() !== "") {
    try {
      parsedReward = JSON.parse(rewardJson);
      if (typeof parsedReward !== 'object' || parsedReward === null || Array.isArray(parsedReward)) {
        throw new Error("獎勵格式不正確，必須是 JSON 物件。");
      }
    } catch (e) {
      throw new Error(`❌ 獎勵 JSON 格式錯誤：${e.message}`);
    }
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const mailSheet = ss.getSheetByName('Mailbox');
  if (!mailSheet) throw new Error("❌ 找不到 Mailbox 工作表。");

  const now = new Date();
  const expiry = parseInt(expiryDays) || 30; // 預設 30 天
  const expiryDate = new Date(now.getTime() + expiry * 24 * 60 * 60 * 1000);

  mailSheet.appendRow([Utilities.getUuid(), title, message, JSON.stringify(parsedReward), now, expiryDate, false, false]);

  return `✅ 郵件 "${title}" 已成功發送！`;
}

/**
 * [管理後台用] 取得所有可作為獎勵的選項 (道具列表等)。
 * @returns {object} 包含所有可選道具的物件。
 */
function getAdminRewardOptions() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const itemSheet = ss.getSheetByName("ItemMaster");
    const achievementSheet = ss.getSheetByName('AchievementLog');
    const fieldMappingSheet = ss.getSheetByName('FieldMapping');

    // 1. 取得所有道具
    const itemData = itemSheet.getDataRange().getValues();
    const itemHeaders = itemData[0];
    const itemIDCol = itemHeaders.indexOf("ItemID");
    const itemNameCol = itemHeaders.indexOf("ItemName");
    const rarityCol = itemHeaders.indexOf("Rarity");

    const items = itemData.slice(1).map(row => ({
        id: row[itemIDCol],
        name: row[itemNameCol] || row[itemIDCol],
        rarity: row[rarityCol] || 0
    })).sort((a, b) => {
        const rarityB = parseInt(b.rarity) || 0;
        const rarityA = parseInt(a.rarity) || 0;
        if (rarityA !== rarityB) return rarityA - rarityB;
        return a.name.localeCompare(b.name);
    });

    // 2. 取得所有可設定的屬性
    // ✅【核心修正】將此處的邏輯與 _getAdminRewardOptions_internal 同步，使用白名單機制。
    const rewardableAttributes = [
        { id: 'Coins', name: '💰 金幣' },
        { id: 'HonorPoints', name: '⭐ 榮譽點數' },
        { id: 'Health', name: '❤️ 健康' },
        { id: 'Mood', name: '😊 心情' },
        { id: 'Energy', name: '⚡ 精力' },
        { id: 'Cleanliness', name: '🧼 清潔度' },
        { id: 'SelfDiscipline', name: '⚖️ 自律' }
    ];

    const mappingData = fieldMappingSheet.getDataRange().getValues();
    const fieldMap = Object.fromEntries(mappingData.slice(1).map(row => [row[1], row[0]]));

    const attributes = rewardableAttributes.map(attr => ({
        id: attr.id,
        name: fieldMap[attr.id] || attr.name // 優先使用 FieldMapping 的名稱，若無則用預設
    })).sort((a, b) => a.name.localeCompare(b.name));


    // 3. 取得待審核成就數量
    let pendingAchievementsCount = 0;
    if (achievementSheet && achievementSheet.getLastRow() > 1) {
        const achData = achievementSheet.getDataRange().getValues();
        const statusCol = achData[0].indexOf('Status');
        if (statusCol !== -1) {
            pendingAchievementsCount = achData.slice(1).filter(row => String(row[statusCol] || '').trim().toLowerCase() === 'pending').length;
        }
    }

    return {
        items: items,
        attributes: attributes, // ✅ 新增：回傳屬性列表
        pendingAchievementsCount: pendingAchievementsCount
    };
}

/**
 * [換裝系統] 取得玩家擁有的所有外觀，以及當前裝備的外觀。
 * @returns {{ownedItems: Array<object>, equippedItems: {background: string, character: string}, counts: object}}
 */
function getWardrobeItems() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const profileSheet = ss.getSheetByName("Profile");
    const wardrobeSheet = ss.getSheetByName("Wardrobe");
    const imageMasterSheet = ss.getSheetByName("ImageMaster");

    if (!profileSheet || !wardrobeSheet || !imageMasterSheet) {
        throw new Error("❌ 找不到必要的資料表 (Profile/Wardrobe/ImageMaster)。");
    }

    // 1. 取得當前裝備
    const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
    const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
    const equippedItems = {
        background: profileRow[profileHeaders.indexOf('currentBackground')] || '',
        character: profileRow[profileHeaders.indexOf('currentCharacter')] || ''
    };

    // 2. 建立 ImageMaster 的快取
    const masterData = imageMasterSheet.getDataRange().getValues();
    const masterHeaders = masterData[0];
    const imageMasterMap = {};
    // ✅ 新增：初始化總數計數器
    const totalCounts = { background: 0, character: 0 };
    masterData.slice(1).forEach(row => {
        const filename = row[masterHeaders.indexOf("Filename")];
        if (filename) {
            imageMasterMap[filename] = {
                ItemName: row[masterHeaders.indexOf("Name")] || filename,
                ItemTypeShort: row[masterHeaders.indexOf("Type")] || 'unknown' // bg 或 char
            };
            // ✅ 新增：累加總數
            if (imageMasterMap[filename].ItemTypeShort === 'bg') {
                totalCounts.background++;
            } else if (imageMasterMap[filename].ItemTypeShort === 'char') {
                totalCounts.character++;
            }
        }
    });

    // 3. 從 Wardrobe 讀取已解鎖的外觀，並加上當前裝備的，最後去重
    const wardrobeData = wardrobeSheet.getDataRange().getValues();
    const wardrobeHeaders = wardrobeData[0];
    const filenameCol = wardrobeHeaders.indexOf("Filename");
    if (filenameCol === -1) throw new Error("❌ 在 'Wardrobe' 工作表中找不到 'Filename' 欄位。");
    const ownedFilenames = new Set(wardrobeData.slice(1).map(row => row[wardrobeHeaders.indexOf("Filename")]));
    ownedFilenames.add(equippedItems.background);
    ownedFilenames.add(equippedItems.character);

    // ✅ 新增：計算已擁有的各類別數量
    const ownedCounts = { background: 0, character: 0 };
    ownedFilenames.forEach(filename => {
        if (imageMasterMap[filename]?.ItemTypeShort === 'bg') ownedCounts.background++;
        else if (imageMasterMap[filename]?.ItemTypeShort === 'char') ownedCounts.character++;
    });
    const ownedItems = Array.from(ownedFilenames).map(filename => {
        const masterInfo = imageMasterMap[filename];
        if (!masterInfo || !filename) return null; // ✅ 增加檢查，避免空檔名

        // ✅【核心修改】將背景和角色的網址產生邏輯分開
        let assetUrl = '';
        if (masterInfo.ItemTypeShort === 'bg') {
            assetUrl = `${getConfig().BG_IMAGE_URL}/${filename}`;
        } else if (masterInfo.ItemTypeShort === 'char') {
            // 您可以在此處為不同玩家或情況設定不同的角色資料夾路徑
            assetUrl = `${getConfig().CHAR_IMAGE_URL}/${filename}`;
        }

        const itemTypeLong = masterInfo.ItemTypeShort === 'bg' ? 'Background' : 'Character';
        return { ItemID: filename, ItemName: masterInfo.ItemName, ItemType: itemTypeLong, AssetUrl: assetUrl };
    }).filter(item => item !== null); // 過濾掉在 ImageMaster 中找不到或檔名為空的項目

    // ✅ 新增：組合最終的計數物件
    const counts = {
        background: { owned: ownedCounts.background, total: totalCounts.background },
        character: { owned: ownedCounts.character, total: totalCounts.character }
    };

    return { ownedItems, equippedItems, counts };
}
/**
 * [換裝系統] 更新玩家裝備中的外觀。
 * @param {{background: string, character: string}} selection - 包含新背景和角色 ItemID 的物件。
 * @returns {object} - 返回最新的玩家狀態，與 getQuickStatus() 格式相同。
 */
function updateEquippedItems(selection) {
  if (!selection || !selection.background || !selection.character) {
    throw new Error("❌ 傳入的選擇無效。");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");
  const headers = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  
  const bgCol = headers.indexOf('currentBackground') + 1;
  const charCol = headers.indexOf('currentCharacter') + 1;

  if (bgCol === 0 || charCol === 0) {
    throw new Error("❌ 在 Profile 工作表中找不到 'currentBackground' 或 'currentCharacter' 欄位。");
  }

  profileSheet.getRange(2, bgCol).setValue(selection.background);
  profileSheet.getRange(2, charCol).setValue(selection.character);

  // 強制將所有待處理的試算表操作完成。
  SpreadsheetApp.flush();

  // ✅【核心修正】直接呼叫 getQuickStatus() 並回傳最新的資料，避免前端讀到舊資料。
  return getQuickStatus();
}

/**
 * [新函式] 處理批次開啟寶箱，包含保底機制。
 * @param {string} itemID - 寶箱的 ItemID。
 * @param {number} quantity - 開啟的數量。
 * @param {boolean} isPurchase - 是否為從商店購買。
 * @returns {Array<object>} - 包含所有獎勵物品的陣列。
 */
function openMultipleChests(itemID, quantity, isPurchase) {
  const ss = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName("Profile");
  const itemSheet = ss.getSheetByName("ItemMaster");
  const inventorySheet = ss.getSheetByName("Inventory");

  // --- 1. 讀取資料 & 檢查費用/數量 ---
  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemRow = itemData.slice(1).find(r => r[itemHeaders.indexOf("ItemID")] === itemID);
  if (!itemRow) throw new Error("❌ 找不到此商品。");

  const item = {};
  itemHeaders.forEach((h, i) => item[h] = itemRow[i]);
  if (item.ItemType !== 'TreasureChest') throw new Error("❌ 此物品不是寶箱。");

  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
  const profile = {};
  profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

  const invData = inventorySheet.getDataRange().getValues();
  const invHeaders = invData[0];

  if (isPurchase) {
    const price = (parseInt(item.BuyPrice) || 0) * quantity;
    const honorPrice = (parseInt(item.HonorBuyPrice) || 0) * quantity;
    if (price > 0 && (parseInt(profile.Coins) || 0) < price) throw new Error(`⚠️ 金幣不足！需要 ${price}。`);
    if (honorPrice > 0 && (parseInt(profile.HonorPoints) || 0) < honorPrice) throw new Error(`⚠️ 榮譽點數不足！需要 ${honorPrice}。`);
    if (price > 0) profile.Coins -= price;
    if (honorPrice > 0) profile.HonorPoints -= honorPrice;
  } else {
    const invIndex = invData.slice(1).findIndex(r => r[invHeaders.indexOf("ItemID")] === itemID);
    if (invIndex === -1) throw new Error("❌ 背包中找不到此寶箱。");
    const currentCount = parseInt(invData[invIndex + 1][invHeaders.indexOf("Count")] || 0);
    if (currentCount < quantity) throw new Error(`⚠️ 寶箱數量不足！需要 ${quantity} 個，但你只有 ${currentCount} 個。`);
    const newCount = currentCount - quantity;
    if (newCount <= 0) inventorySheet.deleteRow(invIndex + 2);
    else inventorySheet.getRange(invIndex + 2, invHeaders.indexOf("Count") + 1).setValue(newCount);
  }

  // --- 2. 執行抽獎 & 保底機制 ---
  const lootTable = JSON.parse(item.LootTableJSON || '[]');
  const wardrobeSheet = ss.getSheetByName("Wardrobe");
  const wardrobeData = wardrobeSheet.getDataRange().getValues();
  const ownedAppearanceFilenames = new Set(wardrobeData.slice(1).map(row => row[wardrobeSheet.getRange(1, 1, 1, wardrobeSheet.getLastColumn()).getValues()[0].indexOf("Filename")]));

  const allResults = [];
  // ✅【核心修改】從 ItemMaster 讀取保底設定，若無則使用預設值
  const pityCounter = parseInt(item.PityCounter) || 10; // 每幾抽保底
  const pityRarity = parseInt(item.PityRarity) || 3;   // 保底幾星或以上

  for (let i = 0; i < quantity; i++) {
    allResults.push(_performSinglePull(lootTable, itemData, itemHeaders, ownedAppearanceFilenames, ss));
  }

  // 執行保底檢查
  for (let i = 0; i < Math.floor(quantity / pityCounter); i++) {
    const batch = allResults.slice(i * pityCounter, (i + 1) * pityCounter);
    const hasPityItem = batch.some(r => r.type === 'item' && parseInt(r.rarity) >= pityRarity);

    if (!hasPityItem) {
      Logger.log(`[多連抽] 第 ${i * pityCounter + 1}-${(i + 1) * pityCounter} 抽未中 ${pityRarity}★ 以上道具，觸發保底！`);
      const guaranteedLootTable = lootTable.filter(l => (l.Type || "ItemRarity") === "ItemRarity" && parseInt(l.Rarity) >= pityRarity);
      if (guaranteedLootTable.length > 0) {
        const guaranteedItem = _performSinglePull(guaranteedLootTable, itemData, itemHeaders, ownedAppearanceFilenames, ss);
        allResults[i * pityCounter + (pityCounter - 1)] = guaranteedItem; // 替換該批次的最後一個結果
        Logger.log(`[多連抽] 保底抽中：${guaranteedItem.rarity}★ ${guaranteedItem.itemName}`);
      }
    }
  }

  // --- 3. 批次結算獎勵 ---
  const currencyRewards = {};
  const itemsToAdd = {};
  const newAppearances = [];
  let duplicateCount = 0;

  allResults.forEach(res => {
    if (res.type === 'currency') currencyRewards[res.rewardID] = (currencyRewards[res.rewardID] || 0) + parseFloat(res.quantity);
    else if (res.type === 'item') itemsToAdd[res.itemID] = (itemsToAdd[res.itemID] || 0) + 1;
    else if (res.type === 'appearance') newAppearances.push(res);
    else if (res.type === 'duplicate_appearance') duplicateCount++;
  });

  const sourceName = `開啟 ${item.ItemName} x${quantity}`;
  applyRewards(profile, currencyRewards, sourceName);
  if (duplicateCount > 0) applyRewards(profile, { 'Coins': 500 * duplicateCount }, `${sourceName} (重複外觀)`);

  if (newAppearances.length > 0) {
    const rows = newAppearances.map(app => [app.filename, app.shortType, new Date()]);
    wardrobeSheet.getRange(wardrobeSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  grantItemsToInventory(itemsToAdd); // 使用批次發放道具函式

  // --- 4. 回傳結果 ---
  return allResults.map(r => ({
    itemName: r.itemName || r.displayName,
    rarity: r.rarity,
    type: r.type,
    rewardID: r.rewardID,
    quantity: r.quantity
  }));
}

/**
 * [全域輔助函式] 執行單次抽獎，並加入抽取外觀的邏輯。
 * @param {Array} lootTable - 獎池設定。
 * @param {Array} allItemData - ItemMaster 的所有資料。
 * @param {Array} allItemHeaders - ItemMaster 的標頭。
 * @param {Set} ownedAppearanceFilenames - 玩家已擁有的外觀檔名 Set。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet 實例。
 * @returns {object} - 抽獎結果物件。
 */
function _performSinglePull(lootTable, allItemData, allItemHeaders, ownedAppearanceFilenames, ss) {
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

    const type = wonLoot.Type || "ItemRarity";

    if (type === "Currency") {
      return { type: 'currency', rewardID: wonLoot.RewardID, quantity: wonLoot.Quantity, displayName: wonLoot.DisplayName, rarity: 'currency' };
    } else if (type === "Appearance") {
      const imageMasterSheet = ss.getSheetByName("ImageMaster");
      const allImageData = imageMasterSheet.getDataRange().getValues();
      const imageMasterHeaders = allImageData[0];
      const appearanceType = wonLoot.AppearanceType; // "Character" or "Background"
      const shortType = appearanceType === 'Background' ? 'bg' : 'char';

      const potentialAppearances = allImageData.slice(1).filter(row => row[imageMasterHeaders.indexOf("Type")] === shortType);
      if (potentialAppearances.length === 0) {
        return { type: 'currency', rewardID: 'Coins', quantity: 500, displayName: '金幣 (補償)', rarity: 'currency' };
      }

      const wonAppearanceRow = potentialAppearances[Math.floor(Math.random() * potentialAppearances.length)];
      const filename = wonAppearanceRow[imageMasterHeaders.indexOf("Filename")];
      const name = wonAppearanceRow[imageMasterHeaders.indexOf("Name")] || filename;

      if (ownedAppearanceFilenames.has(filename)) {
        // ✅【核心修改】抽到重複的
        return { type: 'duplicate_appearance', itemName: name, rarity: 'duplicate' };
      } else {
        // ✅ 抽到新的
        return { type: 'appearance', filename: filename, itemName: name, shortType: shortType, rarity: 'new-appearance' };
      }
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

/**
 * [成就系統] 玩家提交一項新成就以供審核。
 * @param {string} category - 成就類別。
 * @param {string} itemName - 書名或作品名稱。
 * @param {string} description - 成就的詳細說明。
 * @returns {string} 執行結果訊息。
 */
function submitAchievement(category, itemName, description) {
  if (!category || !itemName || !description) {
    throw new Error("❌ 類別、名稱和說明為必填項目。");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let achievementSheet = ss.getSheetByName('AchievementLog');
  if (!achievementSheet) {
    achievementSheet = ss.insertSheet('AchievementLog');
    // ✅ 修正：建立新的標頭，加入 IsLongTerm
    achievementSheet.appendRow(['AchievementID', 'PlayerName', 'Category', 'ItemName', 'Description', 'SubmissionDate', 'Status', 'ApprovalDate', 'RewardJSON', 'RejectionReason', 'IsLongTerm']);
  }

  const profileSheet = ss.getSheetByName('Profile');
  // ✅【修正】從寫死的 'B2' 改為動態尋找 'PlayerName' 欄位，避免欄位順序變動時出錯。
  const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
  const playerNameColIndex = profileHeaders.indexOf('PlayerName');
  
  if (playerNameColIndex === -1) {
    // 如果找不到 PlayerName 欄位，提供一個明確的錯誤訊息
    throw new Error("❌ 在 'Profile' 工作表中找不到 'PlayerName' 欄位標頭。請檢查欄位名稱是否正確。");
  }
  // 從第二列的對應欄位中取得玩家名稱
  const playerName = profileSheet.getRange(2, playerNameColIndex + 1).getValue();

  const achievementID = Utilities.getUuid();
  const submissionDate = new Date();

  achievementSheet.appendRow([achievementID, playerName, category, itemName, description, submissionDate, 'Pending', '', '', '', false]);
  SpreadsheetApp.flush(); // ✅ 關鍵修正：強制將所有待處理的試算表操作完成。
  Utilities.sleep(1500);  // ✅ 終極手段：強制等待 1.5 秒，確保資料已在 Google 伺服器間同步。

  // 發送通知郵件給管理者
  const adminUrl = ScriptApp.getService().getUrl() + `?page=admin`;
  const subject = `[遊戲成就提報] ${playerName} 提交新成就：[${category}] ${itemName}`;
  const body = `
    <h3>您好！</h3>
    <p>玩家 <strong>${playerName}</strong> 提交了一項新完成的成就：</p>
    <p><strong>類別：</strong> ${category}</p>
    <p style="font-size: 18px; font-weight: bold;">名稱/作品： ${itemName}</p>
    <p><strong>說明：</strong></p>
    <p style="padding: 10px; background-color: #f4f4f4; border-radius: 5px; white-space: pre-wrap;">${description}</p> 
    <p>請點擊以下連結，前往管理後台審核此項成就並決定是否發放獎勵：</p>
    <p><a href="${adminUrl}" style="font-size: 16px; padding: 10px 15px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px;">前往管理後台</a></p>
  `;
  MailApp.sendEmail({ to: 'renewmiffy@gmail.com', subject: subject, htmlBody: body }); // 請記得換成您的 Email

  return "✅ 成就已成功提交審核！";
}

/**
 * [管理後台用] 取得所有待審核的成就。
 * @returns {Array<object>} 待審核的成就物件陣列。
 */
function getPendingAchievements() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('AchievementLog');
    if (!sheet) {
      Logger.log('[getPendingAchievements] 錯誤：找不到名為 "AchievementLog" 的工作表。');
      return [];
    }
    if (sheet.getLastRow() < 2) {
      Logger.log('[getPendingAchievements] "AchievementLog" 工作表是空的或只有標頭。');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim()); // ✅【強化】標頭也清理空格
    const statusCol = headers.indexOf('Status');

    if (statusCol === -1) {
      Logger.log(`[getPendingAchievements] 錯誤：在 "AchievementLog" 中找不到 "Status" 欄位標頭。找到的標頭是: [${headers.join(', ')}]`);
      return [];
    }

    const pendingItems = data.slice(1).filter(row => {
      // ✅【強化】比對時，將儲存格內容轉為字串、清理前後空格、並忽略大小寫
      const statusValue = String(row[statusCol] || '').trim();
      return statusValue.toLowerCase() === 'pending';
    }).map(row => {
      const item = {};
      headers.forEach((h, i) => {
        if ((h === 'SubmissionDate' || h === 'ApprovalDate') && row[i] instanceof Date) {
          item[h] = row[i].toISOString();
        } else {
          item[h] = row[i];
        }
      });
      return item;
    });

    Logger.log(`[getPendingAchievements] 查詢完畢。共找到 ${pendingItems.length} 筆待審核的成就。`);
    return pendingItems;
  } catch (e) {
    Logger.log(`[getPendingAchievements] 執行時發生嚴重錯誤: ${e.message}`);
    return []; // 發生任何錯誤都回傳空陣列，避免前端報錯
  }
}

/**
 * [成就審核頁面用] 根據 ID 取得成就詳情。
 * @param {string} achievementId - 成就 ID。
 * @returns {object|null} 成就物件或 null。
 */
function getAchievementDetails(achievementId) {
  if (!achievementId) return null;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('AchievementLog');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('AchievementID');

  const row = data.find(r => r[idCol] === achievementId);
  if (!row) return null;

  const details = {};
  headers.forEach((h, i) => {
    if ((h === 'SubmissionDate' || h === 'ApprovalDate') && row[i] instanceof Date) {
      details[h] = row[i].toISOString();
    } else {
      details[h] = row[i];
    }
  });
  return details;
}

/**
 * [成就審核頁面用] 處理核准或拒絕。
 * @param {string} achievementId - 成就 ID。
 * @param {string} action - 'approved' 或 'rejected'。
 * @param {string} rewardJson - 獎勵的 JSON 字串。
 * @param {string} rejectionReason - 拒絕理由。
 * @param {boolean} isLongTerm - 是否設為長期成就。
 * @returns {string} 結果訊息。
 */
function processAchievement(achievementId, action, rewardJson, rejectionReason, isLongTerm) {
  if (!achievementId) throw new Error("❌ 無效的成就 ID。");
  if (action !== 'approved' && action !== 'rejected') throw new Error("❌ 無效的操作。");

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('AchievementLog');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('AchievementID');
  const statusCol = headers.indexOf('Status');

  const rowIndex = data.findIndex(r => r[idCol] === achievementId);
  if (rowIndex === -1) throw new Error("❌ 找不到此成就紀錄。");
  if (data[rowIndex][statusCol] !== 'Pending') throw new Error("⚠️ 此請求已被處理，請勿重複操作。");

  const newStatus = action === 'approved' ? 'Approved' : 'Rejected';
  sheet.getRange(rowIndex + 1, statusCol + 1).setValue(newStatus);
  sheet.getRange(rowIndex + 1, headers.indexOf('ApprovalDate') + 1).setValue(new Date());
  const achievementItemName = data[rowIndex][headers.indexOf('ItemName')];

  // ✅ 新增：記錄是否為長期成就
  if (action === 'approved') {
    const isLongTermCol = headers.indexOf('IsLongTerm');
    if (isLongTermCol !== -1) {
      sheet.getRange(rowIndex + 1, isLongTermCol + 1).setValue(isLongTerm === true);
    }
    // ✅【新功能】如果設為里程碑，則將其複製到專用的 MilestoneLog 工作表
    if (isLongTerm === true) {
      try {
        let milestoneSheet = ss.getSheetByName('MilestoneLog');
        if (!milestoneSheet) {
          milestoneSheet = ss.insertSheet('MilestoneLog');
          // 只複製前端顯示需要的欄位
          milestoneSheet.appendRow(['AchievementID', 'PlayerName', 'Category', 'ItemName', 'Description', 'ApprovalDate', 'RewardJSON']);
        }
        const achievementData = data[rowIndex];
        const newMilestoneRow = [
          achievementData[headers.indexOf('AchievementID')],
          achievementData[headers.indexOf('PlayerName')],
          achievementData[headers.indexOf('Category')],
          achievementItemName, // ItemName
          achievementData[headers.indexOf('Description')],
          new Date(), // ApprovalDate
          rewardJson // RewardJSON
        ];
        milestoneSheet.appendRow(newMilestoneRow);
      } catch (e) { Logger.log(`將里程碑寫入 MilestoneLog 時出錯: ${e.message}`); }
    }
  }

  if (action === 'approved' && rewardJson && rewardJson.trim() !== '{}' && rewardJson.trim() !== '') {
    // 讀取玩家資料以檢查狀態，判斷是否需要附加警告
    const profileSheet = ss.getSheetByName("Profile");
    const profileHeaders = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
    const profileRow = profileSheet.getRange(2, 1, 1, profileHeaders.length).getValues()[0];
    const profile = {};
    profileHeaders.forEach((k, i) => profile[k] = profileRow[i]);

    // 計算當前效果
    const activeStatuses = evaluateStatusRules(profile);
    const effectsSummary = calculateEffectsSummary(activeStatuses);
    
    let mailMessage = `恭喜你完成了成就「${achievementItemName}」！<br><br>這是給你的獎勵，請查收。`;

    // 檢查是否有懲罰，並附加警告訊息
    if (effectsSummary.globalRewardModifier < 1) {
      const penaltyPercent = Math.round((1 - effectsSummary.globalRewardModifier) * 100);
      const penaltyReasonStatus = activeStatuses.find(s => s.GlobalRewardModifier && parseFloat(s.GlobalRewardModifier) < 1);
      let reasonText = penaltyReasonStatus ? `因為你現在處於「${penaltyReasonStatus.狀態名稱}」狀態` : "因為你目前的狀態不佳";
      
      mailMessage += `<br><br><div style="color: #721c24; border: 1px solid #f5c6cb; background-color: #f8d7da; padding: 10px; border-radius: 5px;"><strong>⚠️ 注意：</strong>${reasonText}，所有獎勵將會減少 <strong>${penaltyPercent}%</strong>。要好好照顧自己喔！</div>`;
    }

    const mailTitle = `成就獎勵：${achievementItemName}`;
    sendAdminMail(mailTitle, mailMessage, rewardJson, 30); // 發送獎勵郵件
    sheet.getRange(rowIndex + 1, headers.indexOf('RewardJSON') + 1).setValue(rewardJson);
    return `✅ 成就已核准，並已發送獎勵郵件！`;
  }

  // ✅ 新增：如果拒絕，發送通知郵件
  if (action === 'rejected' && rejectionReason) {
    const reasonCol = headers.indexOf('RejectionReason');
    if (reasonCol !== -1) {
      sheet.getRange(rowIndex + 1, reasonCol + 1).setValue(rejectionReason);
    }
    const mailTitle = `成就審核結果通知`;
    const mailMessage = `很遺憾，您提交的成就「${achievementItemName}」未獲核准。\n\n管理員給您的建議：\n${rejectionReason}\n\n請再接再厲！`;
    sendAdminMail(mailTitle, mailMessage, "{}", 30); // 發送不含獎勵的通知郵件
    return `✅ 成就已拒絕，並已發送附帶理由的通知郵件。`;
  }

  return `✅ 成就已標示為 [${newStatus}]。`;
}

/**
 * [新函式][內部用] 取得管理後台獎勵選項 (優化版)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 已開啟的 Spreadsheet 物件。
 */
function _getAdminRewardOptions_internal(ss) {
  const itemSheet = ss.getSheetByName("ItemMaster");
  const achievementSheet = ss.getSheetByName('AchievementLog');
  const fieldMappingSheet = ss.getSheetByName('FieldMapping');

  // 1. 取得所有道具
  const itemData = itemSheet.getDataRange().getValues();
  const itemHeaders = itemData[0];
  const itemIDCol = itemHeaders.indexOf("ItemID");
  const itemNameCol = itemHeaders.indexOf("ItemName");
  const rarityCol = itemHeaders.indexOf("Rarity");

  const items = itemData.slice(1).map(row => ({
    id: row[itemIDCol],
    name: row[itemNameCol] || row[itemIDCol],
    rarity: row[rarityCol] || 0
  })).sort((a, b) => (parseInt(a.rarity) || 0) - (parseInt(b.rarity) || 0) || a.name.localeCompare(b.name));

  // 2. 取得所有可設定的屬性
  // ✅【核心修正】不再讀取整個 FieldMapping，而是定義一個可作為獎勵的屬性白名單。
  // 這樣可以避免將不應修改的欄位（如玩家名稱）顯示在獎勵編輯器中。
  const rewardableAttributes = [
    { id: 'Coins', name: '💰 金幣' },
    { id: 'HonorPoints', name: '⭐ 榮譽點數' },
    { id: 'Health', name: '❤️ 健康' },
    { id: 'Mood', name: '😊 心情' },
    { id: 'Energy', name: '⚡ 精力' },
    { id: 'Cleanliness', name: '🧼 清潔度' },
    { id: 'SelfDiscipline', name: '⚖️ 自律' }
  ];

  // 為了保持彈性，我們仍然可以從 FieldMapping 讀取中文名稱，但只讀取白名單中的屬性。
  const mappingData = fieldMappingSheet.getDataRange().getValues();
  const fieldMap = Object.fromEntries(mappingData.slice(1).map(row => [row[1], row[0]]));

  const attributes = rewardableAttributes.map(attr => ({
      id: attr.id,
      name: fieldMap[attr.id] || attr.name // 優先使用 FieldMapping 的名稱，若無則用預設
  })).sort((a, b) => a.name.localeCompare(b.name));

  // 3. 取得待審核成就數量 (這部分邏輯與原函式相同，可以保持)
  let pendingAchievementsCount = 0;
  if (achievementSheet && achievementSheet.getLastRow() > 1) {
    const achData = achievementSheet.getDataRange().getValues();
    const statusCol = achData[0].indexOf('Status');
    if (statusCol !== -1) pendingAchievementsCount = achData.slice(1).filter(row => String(row[statusCol] || '').trim().toLowerCase() === 'pending').length;
  }

  return { items, attributes, pendingAchievementsCount };
}

/**
 * [玩家用] 取得已核准的成就歷史紀錄，分為「里程碑」和「近期成果」。
 * @returns {{milestones: Array<object>, recents: Array<object>}} 包含兩類成就的物件。
 */
function getCompletedAchievements() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('AchievementLog');
    const milestoneSheet = ss.getSheetByName('MilestoneLog');

    let milestones = [];
    let recents = [];

    // 1. 【效能優化】直接從 MilestoneLog 工作表讀取里程碑
    if (milestoneSheet && milestoneSheet.getLastRow() >= 2) {
      Logger.log('[getCompletedAchievements] 正在從 MilestoneLog 工作表讀取里程碑...');
      const milestoneData = milestoneSheet.getDataRange().getValues();
      const milestoneHeaders = milestoneData[0].map(h => String(h).trim());
      milestones = milestoneData.slice(1).map(row => {
        const item = {};
        milestoneHeaders.forEach((h, i) => {
          item[h] = (h === 'ApprovalDate') && row[i] instanceof Date ? row[i].toISOString() : row[i];
        });
        return item;
      }).sort((a, b) => new Date(b.ApprovalDate) - new Date(a.ApprovalDate));
    }

    // 2. 從 AchievementLog 工作表讀取近期成果
    if (logSheet && logSheet.getLastRow() >= 2) {
      const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const statusCol = headers.indexOf('Status');
      const isLongTermCol = headers.indexOf('IsLongTerm');

      if (statusCol !== -1) {
        const lastRow = logSheet.getLastRow();
        const startRow = Math.max(2, lastRow - 199);
        const numRows = lastRow - startRow + 1;
        const recentData = logSheet.getRange(startRow, 1, numRows, headers.length).getValues();

        const recentRows = recentData.filter(row => {
          const isApproved = String(row[statusCol] || '').trim().toLowerCase() === 'approved';
          const isLongTerm = (isLongTermCol !== -1) ? String(row[isLongTermCol]).toLowerCase() === 'true' : false;
          return isApproved && !isLongTerm;
        });

        recents = recentRows.map(row => {
          const item = {};
          headers.forEach((h, i) => {
            item[h] = (h === 'SubmissionDate' || h === 'ApprovalDate') && row[i] instanceof Date ? row[i].toISOString() : row[i];
          });
          return item;
        }).sort((a, b) => new Date(b.ApprovalDate) - new Date(a.ApprovalDate)).slice(0, 50);
      }
    }

    return { milestones: milestones, recents: recents };
  } catch (e) {
    Logger.log(`[getCompletedAchievements] 執行時發生嚴重錯誤: ${e.message}`);
    return { milestones: [], recents: [] }; // 發生任何錯誤都回傳空陣列，避免前端報錯
  }
}
/**
 * [新函式][管理後台用] 取得任務管理頁面需要的所有資料。
 * @returns {object} 包含所有任務列表和獎勵選項的物件。
 */
function getTaskManagementData() {
  try {
    // 【效能優化】只打開一次試算表
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 【效能優化】呼叫新的內部函式，並傳入已開啟的 ss 物件
    const dailyTasks = _getDailyTaskList_internal(ss);
    const learningSkills = _getAllSkills_internal(ss);
    const missions = _getMissionList_internal(ss);
    const rewardOptions = _getAdminRewardOptions_internal(ss);

    return {
      dailyTasks: dailyTasks,
      learningSkills: learningSkills,
      missions: missions,
      rewardOptions: rewardOptions
    };
  } catch (e) {
    Logger.log(`[getTaskManagementData] 執行時發生錯誤: ${e.stack}`);
    return { error: e.message };
  }
}

/**
 * [新函式][內部用] 取得日常任務列表 (優化版)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 已開啟的 Spreadsheet 物件。
 */
function _getDailyTaskList_internal(ss) {
  const sheet = ss.getSheetByName("DailyTasks");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");

  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    let lastDateStr = (obj.LastDoneDate instanceof Date) ? Utilities.formatDate(obj.LastDoneDate, Session.getScriptTimeZone(), "yyyy/MM/dd") : "";
    return {
      id: obj.TaskID,
      name: obj.任務名稱,
      effects: obj.Effects,
      fulfilledToday: (lastDateStr === todayStr)
    };
  });
}

/**
 * [新函式][內部用] 取得所有技能列表 (優化版)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 已開啟的 Spreadsheet 物件。
 */
function _getAllSkills_internal(ss) {
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

/**
 * [新函式][內部用] 取得任務中心列表 (優化版)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 已開啟的 Spreadsheet 物件。
 */
function _getMissionList_internal(ss) {
  // 為了避免修改複雜的 getMissionList 函式，我們暫時讓它保持原樣
  // 這裡的呼叫仍然會多開一次試算表，但已從 4 次減少到 2 次，效能會有顯著提升
  // 未來可以進一步將 getMissionList 內部邏輯也拆分出來
  return getMissionList();
}

/**
 * [新函式][管理後台用] 儲存單一任務的設定變更。
 * @param {string} taskType - 任務類型 ('daily', 'skill', 'mission')。
 * @param {string} taskId - 要更新的任務 ID。
 * @param {object} updatedData - 一個包含要更新欄位和新值的物件。例如：{ "任務名稱": "新的名稱", "Effects": "{...}" }
 * @returns {string} 執行結果訊息。
 */
function saveTaskConfiguration(taskType, taskId, updatedData) {
    if (!taskType || !taskId || !updatedData) {
        throw new Error("❌ 缺少必要的參數 (taskType, taskId, updatedData)。");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet;
    let idColumnName;

    // 根據任務類型選擇對應的工作表和 ID 欄位名稱
    switch (taskType) {
        case 'daily':
            sheet = ss.getSheetByName("DailyTasks");
            idColumnName = "TaskID";
            break;
        case 'skill':
            sheet = ss.getSheetByName("SkillMaster");
            idColumnName = "SkillID";
            break;
        case 'mission':
            sheet = ss.getSheetByName("MissionCenter");
            idColumnName = "MissionID";
            break;
        default:
            throw new Error(`❌ 無效的任務類型: ${taskType}`);
    }

    if (!sheet) throw new Error(`❌ 找不到工作表: ${sheet.getName()}`);

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf(idColumnName);

    if (idColIndex === -1) throw new Error(`❌ 在 ${sheet.getName()} 中找不到 ID 欄位 "${idColumnName}"。`);

    // 找到要更新的列
    const rowIndex = data.findIndex(row => row[idColIndex] === taskId);
    if (rowIndex === -1) throw new Error(`❌ 在 ${sheet.getName()} 中找不到 ID 為 "" 的任務。`);

    // 逐一更新指定的欄位
    Object.entries(updatedData).forEach(([headerName, newValue]) => {
        const colIndex = headers.indexOf(headerName);
        if (colIndex !== -1) {
            sheet.getRange(rowIndex + 1, colIndex + 1).setValue(newValue);
        }
    });

    return `✅ 已成功更新 ${taskType} 任務 [${taskId}] 的設定。`;
}

/**
 * [新函式][管理後台用] 刪除一個指定的任務。
 * @param {string} taskType - 任務類型 ('daily', 'skill', 'mission')。
 * @param {string} taskId - 要刪除的任務 ID。
 * @returns {string} 執行結果訊息。
 */
function deleteTask(taskType, taskId) {
    if (!taskType || !taskId) {
        throw new Error("❌ 缺少必要的參數 (taskType, taskId)。");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet;
    let idColumnName;

    switch (taskType) {
        case 'daily': sheet = ss.getSheetByName("DailyTasks"); idColumnName = "TaskID"; break;
        case 'skill': sheet = ss.getSheetByName("SkillMaster"); idColumnName = "SkillID"; break;
        case 'mission': sheet = ss.getSheetByName("MissionCenter"); idColumnName = "MissionID"; break;
        default: throw new Error(`❌ 無效的任務類型: ${taskType}`);
    }

    if (!sheet) throw new Error(`❌ 找不到工作表: ${sheet.getName()}`);

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf(idColumnName);

    if (idColIndex === -1) throw new Error(`❌ 在 ${sheet.getName()} 中找不到 ID 欄位 "${idColumnName}"。`);

    // 從後往前找，這樣刪除時才不會影響到尚未檢查的列的索引
    for (let i = data.length - 1; i > 0; i--) {
        if (data[i][idColIndex] === taskId) {
            sheet.deleteRow(i + 1); // i 是 data 陣列的索引 (0-based)，工作表列號是 1-based
            return `✅ 已成功刪除 ${taskType} 任務 [${taskId}]。`;
        }
    }

    throw new Error(`❌ 在 ${sheet.getName()} 中找不到 ID 為 "${taskId}" 的任務可供刪除。`);
}

/**
 * [新函式][管理後台用] 新增一個任務。
 * @param {string} taskType - 任務類型 ('daily', 'skill', 'mission')。
 * @param {object} taskData - 新任務的資料，至少需要包含 ID 和名稱。例如 { "TaskID": "D010", "任務名稱": "新任務" }
 * @returns {string} 執行結果訊息。
 */
function createTask(taskType, taskData) {
    if (!taskType || !taskData) {
        throw new Error("❌ 缺少必要的參數 (taskType, taskData)。");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet, idColumnName;

    switch (taskType) {
        case 'daily': sheet = ss.getSheetByName("DailyTasks"); idColumnName = "TaskID"; break;
        case 'skill': sheet = ss.getSheetByName("SkillMaster"); idColumnName = "SkillID"; break;
        case 'mission': sheet = ss.getSheetByName("MissionCenter"); idColumnName = "MissionID"; break;
        default: throw new Error(`❌ 無效的任務類型: ${taskType}`);
    }

    if (!sheet) throw new Error(`❌ 找不到工作表: ${sheet.getName()}`);

    const newId = taskData[idColumnName];

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf(idColumnName);

    // ✅【核心修改】如果前端沒有提供 ID，則自動產生一個。
    if (!newId || String(newId).trim() === '') {
        let prefix = '';
        if (taskType === 'daily') {
            prefix = 'T'; // 日常任務 (DailyTasks)
        } else if (taskType === 'skill') {
            prefix = 'S'; // 學習技能 (SkillMaster)
        } else if (taskType === 'mission') {
            // 根據前端傳來的「類型」決定前綴
            const missionType = taskData['類型'];
            prefix = (missionType === 'Daily') ? 'D' : 'A'; // D for Daily mission, A for Achievement mission
        }
        if (!prefix) throw new Error(`❌ 無法為類型 "${taskType}" (子類型: ${taskData['類型']}) 產生 ID 前綴。`);

        const existingIds = data.slice(1).map(row => String(row[idColIndex] || '').trim());
        let maxNum = 0;
        existingIds.forEach(id => {
            if (id.startsWith(prefix)) {
                const num = parseInt(id.substring(1), 10);
                if (!isNaN(num) && num > maxNum) {
                    maxNum = num;
                }
            }
        });
        // 產生新 ID，例如 D004, S011
        const generatedId = prefix + String(maxNum + 1).padStart(3, '0');
        taskData[idColumnName] = generatedId; // 將產生的 ID 加回要寫入的資料中
    }

    // 建立新的一列
    const newRow = headers.map(header => {
        if (taskData.hasOwnProperty(header)) return taskData[header];
        // 為關鍵欄位提供預設值
        if (header === 'Effects' || header === '獎勵') return '{}';
        if (header === 'TotalDoneCount' || header === 'StreakCount') return 0;
        return ''; // 其他欄位預設為空
    });

    sheet.appendRow(newRow);

    return `✅ 已成功在 ${sheet.getName()} 中新增任務 [${taskData[idColumnName]}]。`;
}
