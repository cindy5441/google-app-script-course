// ============================================================
// Session 5：資料結構化與觸發器
// 日期：115/05/16（六）09:00~12:00
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. 基本語法延伸（function & return）
//   2. 資料讀取與結構化 (Array)
//   3. 觸發器應用（定期更新資料）
//   4. 實作：定時抓取表格並生成清單
// ============================================================

// ============================================================
// 第一部分：函數進階 — function & return
// ============================================================

/**
 * 多回傳值的函數設計
 * 說明：用物件 (Object) 或陣列回傳多個值
 */
function 函數進階示範() {
  // --- 回傳物件 ---
  var 員工資訊 = 取得員工統計();
  Logger.log("總人數：" + 員工資訊.總人數);
  Logger.log("平均薪資：" + 員工資訊.平均薪資);
  Logger.log("最高薪資：" + 員工資訊.最高薪資);

  // --- 函數作為參數（回呼函數）---
  var 數字 = [3, 1, 4, 1, 5, 9, 2, 6];
  var 篩選結果 = 篩選陣列(數字, function(x) {
    return x > 3;  // 只保留大於 3 的數字
  });
  Logger.log("大於 3 的數字：" + 篩選結果); // [4, 5, 9, 6]
}

/**
 * 從試算表讀取員工資料並計算統計值
 * @returns {Object} 包含統計資訊的物件
 */
function 取得員工統計() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("專案人員");
  if (!sheet) {
    return { 總人數: 0, 平均薪資: 0, 最高薪資: 0 };
  }

  var 資料 = sheet.getDataRange().getValues();
  var 薪資列表 = [];

  for (var i = 1; i < 資料.length; i++) {
    薪資列表.push(資料[i][3]); // D 欄：月薪
  }

  // 計算統計值
  var 總人數 = 薪資列表.length;
  var 薪資總和 = 薪資列表.reduce(function(sum, val) { return sum + val; }, 0);
  var 平均薪資 = Math.round(薪資總和 / 總人數);
  var 最高薪資 = Math.max.apply(null, 薪資列表);
  var 最低薪資 = Math.min.apply(null, 薪資列表);

  // 回傳一個物件（包含多個值）
  return {
    總人數: 總人數,
    平均薪資: 平均薪資,
    最高薪資: 最高薪資,
    最低薪資: 最低薪資,
    薪資總和: 薪資總和
  };
}

/**
 * 篩選陣列的通用函數
 * @param {Array} 陣列 - 要篩選的陣列
 * @param {Function} 條件函數 - 篩選條件（回傳 true/false）
 * @returns {Array} 篩選後的陣列
 */
function 篩選陣列(陣列, 條件函數) {
  var 結果 = [];
  for (var i = 0; i < 陣列.length; i++) {
    if (條件函數(陣列[i])) {
      結果.push(陣列[i]);
    }
  }
  return 結果;
}

// ============================================================
// 第二部分：資料結構化
// ============================================================

/**
 * 將試算表資料轉換為結構化物件陣列
 * 說明：比二維陣列更直覺、更好維護
 */
function 資料結構化示範() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("專案人員");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 請先執行「初始化專案資料」");
    return;
  }

  var 資料 = sheet.getDataRange().getValues();
  var 標題列 = 資料[0]; // ["姓名", "部門", "職稱", "月薪", ...]

  // 將二維陣列轉為物件陣列
  var 員工列表 = [];
  for (var i = 1; i < 資料.length; i++) {
    var 員工 = {};
    for (var j = 0; j < 標題列.length; j++) {
      員工[標題列[j]] = 資料[i][j];
    }
    員工列表.push(員工);
  }

  // 現在可以用 .屬性名 來存取，更直覺！
  Logger.log("===== 結構化資料 =====");
  for (var k = 0; k < 員工列表.length; k++) {
    Logger.log(員工列表[k]["姓名"] + " - " + 員工列表[k]["部門"] + " - NT$" + 員工列表[k]["月薪"]);
  }

  // 進階：依部門分組
  var 部門分組 = 依欄位分組(員工列表, "部門");

  Logger.log("\n===== 部門分組 =====");
  for (var 部門 in 部門分組) {
    Logger.log(部門 + "：" + 部門分組[部門].length + " 人");
    for (var m = 0; m < 部門分組[部門].length; m++) {
      Logger.log("  - " + 部門分組[部門][m]["姓名"]);
    }
  }
}

/**
 * 依指定欄位將資料分組
 * @param {Array<Object>} 資料 - 物件陣列
 * @param {string} 欄位名 - 要分組的欄位名稱
 * @returns {Object} 分組結果
 */
function 依欄位分組(資料, 欄位名) {
  var 分組 = {};
  for (var i = 0; i < 資料.length; i++) {
    var 值 = 資料[i][欄位名];
    if (!分組[值]) {
      分組[值] = [];
    }
    分組[值].push(資料[i]);
  }
  return 分組;
}

// ============================================================
// 第三部分：觸發器應用 — 定期更新資料
// ============================================================

/**
 * 設定每小時自動更新的觸發器
 */
function 設定定期更新觸發器() {
  // 清除同名舊觸發器
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "定時更新專案狀態") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 建立每小時觸發器
  ScriptApp.newTrigger("定時更新專案狀態")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("✅ 每小時更新觸發器已設定");
  SpreadsheetApp.getUi().alert("✅ 已設定每小時自動更新！");
}

/**
 * 定時更新專案狀態
 * 說明：自動計算專案進度、逾期檢查，並生成摘要、標記顏色及自動通知（包含加班預警）
 */
function 定時更新專案狀態() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 專案表 = ss.getSheetByName("專案追蹤");
    if (!專案表) return;

    var 資料 = 專案表.getDataRange().getValues();
    var 今天 = new Date();
    var 更新結果 = [];
    var 背景顏色 = [];
    var 逾期專案 = [];
    var 員工任務數 = {};

    for (var i = 1; i < 資料.length; i++) {
      var 專案名稱 = 資料[i][0];
      var 負責人 = 資料[i][1];
      var 截止日 = new Date(資料[i][4]); // E 欄：截止日期
      var 進度 = 資料[i][5];             // F 欄：完成進度
      var 狀態;
      var 顏色 = "#ffffff"; // 預設白色

      // 累計員工進行中的任務數 (進度小於 100 且有負責人)
      if (負責人 && 負責人.toString().trim() !== "" && 進度 < 100) {
        員工任務數[負責人] = (員工任務數[負責人] || 0) + 1;
      }

      if (進度 >= 100) {
        狀態 = "✅ 已完成";
        顏色 = "#e8f5e9"; // 淺綠色
      } else if (今天 > 截止日) {
        狀態 = "🔴 已逾期";
        顏色 = "#ffebee"; // 淺紅色
        逾期專案.push("專案：" + 專案名稱 + " (負責人：" + 負責人 + ")");
      } else {
        var 剩餘天數 = Math.ceil((截止日 - 今天) / (1000 * 60 * 60 * 24));
        if (剩餘天數 <= 3) {
          狀態 = "🟡 即將到期 (" + 剩餘天數 + "天)";
          顏色 = "#fff3e0"; // 淺橘色
        } else {
          狀態 = "🟢 進行中 (" + 剩餘天數 + "天)";
        }
      }

      更新結果.push([狀態]);
      
      // 準備整列的背景顏色陣列 (A到G共7欄)
      var 列顏色 = [];
      for(var c = 0; c < 7; c++) 列顏色.push(顏色);
      背景顏色.push(列顏色);
    }

    // 批次寫入狀態與顏色
    if (更新結果.length > 0) {
      專案表.getRange(2, 7, 更新結果.length, 1).setValues(更新結果);
      專案表.getRange(2, 1, 背景顏色.length, 7).setBackgrounds(背景顏色);
    }

    // 檢查員工負擔 (假設超過 4 件進行中專案為過重)
    var 超載員工 = [];
    var 負擔上限 = 4;
    for (var 員工 in 員工任務數) {
      if (員工任務數[員工] > 負擔上限) {
        超載員工.push(員工 + " (進行中：" + 員工任務數[員工] + " 件)");
      }
    }

    // 整合發送通知信 (逾期或超載)
    var 使用者信箱 = Session.getActiveUser().getEmail();
    if (使用者信箱 && (逾期專案.length > 0 || 超載員工.length > 0)) {
      var 信件標題 = "⚠️ 專案管理系統預警通知";
      var 信件內容 = "您好，\n\n系統偵測到以下需要注意的事項：\n\n";
      
      if (逾期專案.length > 0) {
        信件內容 += "【逾期專案】\n" + 逾期專案.join("\n") + "\n\n";
      }
      
      if (超載員工.length > 0) {
        信件內容 += "【員工負擔過重預警】\n" + 超載員工.join("\n") + "\n請考慮重新分配任務以防加班。\n\n";
      }
      
      信件內容 += "請前往試算表查看詳細狀況。\n\n系統自動發送";
      
      MailApp.sendEmail(使用者信箱, 信件標題, 信件內容);
      Logger.log("📧 已發送預警通知信給：" + 使用者信箱);
    }

    // 新增：每日一次的指定信箱逾期提醒
    if (逾期專案.length > 0) {
      var properties = PropertiesService.getScriptProperties();
      var todayStr = Utilities.formatDate(今天, "Asia/Taipei", "yyyy-MM-dd");
      var lastSentDate = properties.getProperty("LAST_OVERDUE_EMAIL_DATE");
      
      if (lastSentDate !== todayStr) {
        var 指定信箱 = "sweewcindy@gmail.com";
        var 信件標題 = "📌 每日逾期任務提醒";
        var 提醒內容 = "您好，\n\n系統偵測到目前仍有逾期未處理的任務。請協助提醒以下員工：\n\n";
        
        for(var u = 0; u < 逾期專案.length; u++) {
          提醒內容 += "- " + 逾期專案[u] + "\n";
        }
        
        提醒內容 += "\n請登入試算表確認進度。\n\n系統自動發送";
        
        MailApp.sendEmail(指定信箱, 信件標題, 提醒內容);
        Logger.log("📧 已發送每日逾期提醒信至：" + 指定信箱);
        
        // 紀錄今天已發送
        properties.setProperty("LAST_OVERDUE_EMAIL_DATE", todayStr);
      }
    }

    // 更新時間戳記
    專案表.getRange("I1").setValue("最後更新");
    專案表.getRange("I2").setValue(Utilities.formatDate(今天, "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    Logger.log("✅ 專案狀態已更新（" + 更新結果.length + " 筆）");

  } catch (錯誤) {
    Logger.log("❌ 定時更新錯誤：" + 錯誤.message);
  }
}

/**
 * 彈窗提醒逾期任務
 * 說明：掃描專案狀態，如果有逾期任務，彈出對話框提醒使用者
 */
function 彈窗提醒逾期任務() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 專案表 = ss.getSheetByName("專案追蹤");
    if (!專案表) return;

    var 資料 = 專案表.getDataRange().getValues();
    var 逾期專案 = [];

    for (var i = 1; i < 資料.length; i++) {
      var 狀態 = String(資料[i][6]); // G 欄
      if (狀態.indexOf("已逾期") >= 0) {
        逾期專案.push("- " + 資料[i][0] + " (負責人：" + 資料[i][1] + ")");
      }
    }

    if (逾期專案.length > 0) {
      var 訊息 = "⚠️ 注意！發現以下逾期任務：\n\n" + 逾期專案.join("\n") + "\n\n請盡速處理！";
      SpreadsheetApp.getUi().alert("逾期任務提醒", 訊息, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert("通知", "🎉 目前沒有逾期任務，太棒了！", SpreadsheetApp.getUi().ButtonSet.OK);
    }

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第四部分：實作 — 定時抓取表格並生成清單
// ============================================================

/**
 * 生成部門人員清單
 * 說明：從專案人員表讀取資料，自動產生按部門分類的人員清單
 */
function 生成部門清單() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 人員表 = ss.getSheetByName("專案人員");
    if (!人員表) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化專案資料」");
      return;
    }

    // 讀取並結構化資料
    var 資料 = 人員表.getDataRange().getValues();
    var 標題 = 資料[0];
    var 員工 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) {
        obj[標題[j]] = 資料[i][j];
      }
      員工.push(obj);
    }

    // 依部門分組
    var 部門 = 依欄位分組(員工, "部門");

    // 建立或清除「部門清單」工作表
    var 清單表 = ss.getSheetByName("部門清單");
    if (清單表) {
      清單表.clear();
    } else {
      清單表 = ss.insertSheet("部門清單");
    }

    // 寫入標題
    清單表.getRange("A1").setValue("📋 部門人員清單");
    清單表.getRange("A1").setFontSize(16).setFontWeight("bold");
    清單表.getRange("A2").setValue("更新時間：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    var 目前列 = 4;

    // 遍歷每個部門
    for (var 部門名 in 部門) {
      var 成員 = 部門[部門名];

      // 部門標題
      清單表.getRange(目前列, 1).setValue("🏢 " + 部門名 + " (" + 成員.length + " 人)");
      清單表.getRange(目前列, 1, 1, 4).merge();
      清單表.getRange(目前列, 1).setFontSize(13).setFontWeight("bold").setBackground("#e3f2fd");
      目前列++;

      // 欄位標題
      清單表.getRange(目前列, 1, 1, 4).setValues([["姓名", "職稱", "月薪", "Email"]]);
      清單表.getRange(目前列, 1, 1, 4).setFontWeight("bold").setBackground("#bbdefb");
      目前列++;

      // 成員資料
      for (var n = 0; n < 成員.length; n++) {
        清單表.getRange(目前列, 1, 1, 4).setValues([
          [成員[n]["姓名"], 成員[n]["職稱"], 成員[n]["月薪"], 成員[n]["Email"]]
        ]);
        清單表.getRange(目前列, 3).setNumberFormat("#,##0");
        目前列++;
      }

      目前列++; // 空一列
    }

    // 調整欄寬
    for (var c = 1; c <= 4; c++) {
      清單表.autoResizeColumn(c);
    }

    Logger.log("✅ 部門清單已生成！");
    SpreadsheetApp.getUi().alert("✅ 部門清單已生成！請查看「部門清單」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 生成專案進度摘要
 */
function 生成專案摘要() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 專案表 = ss.getSheetByName("專案追蹤");
    if (!專案表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「專案追蹤」工作表");
      return;
    }

    // 先更新狀態
    定時更新專案狀態();

    var 資料 = 專案表.getDataRange().getValues();
    var 標題 = 資料[0];
    var 專案列表 = [];
    
    // 1. 將資料結構化
    for (var i = 1; i < 資料.length; i++) {
      var 專案 = {};
      for (var j = 0; j < 標題.length; j++) {
        專案[標題[j]] = 資料[i][j];
      }
      專案列表.push(專案);
    }

    // 2. 依優先級分組
    var 優先級分組 = 依欄位分組(專案列表, "優先級");

    // 3. 建立摘要字串
    var 摘要 = "📊 專案優先級摘要\n\n";
    摘要 += "總專案數：" + 專案列表.length + "\n\n";

    var 順序 = ["高", "中", "低"];
    for (var k = 0; k < 順序.length; k++) {
      var 級別 = 順序[k];
      var 該級別專案 = 優先級分組[級別] || [];
      摘要 += "🔥 優先級 [" + 級別 + "]：" + 該級別專案.length + " 件\n";
      for (var m = 0; m < 該級別專案.length; m++) {
        摘要 += "  - " + 該級別專案[m]["專案名稱"] + " (" + 該級別專案[m]["狀態"] + ")\n";
      }
      摘要 += "\n";
    }

    SpreadsheetApp.getUi().alert(摘要);
    Logger.log(摘要);

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 生成部門薪資報表
 * 說明：自動計算每個部門的平均薪資與人數，並產出至「薪資報表」工作表
 */
function 生成薪資報表() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 人員表 = ss.getSheetByName("專案人員");
    if (!人員表) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化專案資料」");
      return;
    }

    // 1. 讀取並結構化資料
    var 資料 = 人員表.getDataRange().getValues();
    var 標題 = 資料[0];
    var 員工列表 = [];
    for (var i = 1; i < 資料.length; i++) {
      var 員工 = {};
      for (var j = 0; j < 標題.length; j++) {
        員工[標題[j]] = 資料[i][j];
      }
      員工列表.push(員工);
    }

    // 2. 依部門分組
    var 部門分組 = 依欄位分組(員工列表, "部門");

    // 3. 計算各部門平均薪資與人數
    var 報表資料 = [];
    for (var 部門名 in 部門分組) {
      var 成員 = 部門分組[部門名];
      var 人數 = 成員.length;
      var 總薪資 = 0;
      for (var k = 0; k < 人數; k++) {
        總薪資 += Number(成員[k]["月薪"]) || 0;
      }
      var 平均薪資 = Math.round(總薪資 / 人數);
      報表資料.push([部門名, 人數, 平均薪資, 總薪資]);
    }

    // 4. 建立或清除「薪資報表」工作表
    var 報表表 = ss.getSheetByName("薪資報表");
    if (報表表) {
      報表表.clear();
    } else {
      報表表 = ss.insertSheet("薪資報表");
    }

    // 5. 寫入標題
    報表表.getRange("A1").setValue("💰 部門薪資報表");
    報表表.getRange("A1").setFontSize(16).setFontWeight("bold");
    報表表.getRange("A2").setValue("更新時間：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    var 標題列 = [["部門", "人數", "平均月薪", "總月薪"]];
    報表表.getRange(4, 1, 1, 4).setValues(標題列)
      .setFontWeight("bold")
      .setBackground("#e8f5e9");

    // 6. 寫入資料
    if (報表資料.length > 0) {
      報表表.getRange(5, 1, 報表資料.length, 4).setValues(報表資料);
      // 設定數字格式
      報表表.getRange(5, 3, 報表資料.length, 2).setNumberFormat("#,##0");
    }

    // 7. 調整欄寬
    for (var c = 1; c <= 4; c++) {
      報表表.autoResizeColumn(c);
    }

    Logger.log("✅ 薪資報表已生成！");
    SpreadsheetApp.getUi().alert("✅ 薪資報表已生成！請查看「薪資報表」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 建立任務看板
 * 說明：將專案資料結構化並依狀態分組，產生 Kanban 看板視圖
 */
function 生成任務看板() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 專案表 = ss.getSheetByName("專案追蹤");
    if (!專案表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「專案追蹤」工作表");
      return;
    }

    // 先更新狀態確保資料最新
    定時更新專案狀態();

    // 1. 讀取並結構化資料
    var 資料 = 專案表.getDataRange().getValues();
    var 標題 = 資料[0];
    var 專案列表 = [];
    for (var i = 1; i < 資料.length; i++) {
      var 專案 = {};
      for (var j = 0; j < 標題.length; j++) {
        專案[標題[j]] = 資料[i][j];
      }
      專案列表.push(專案);
    }

    // 2. 依狀態基礎分類分組
    var 看板分類 = {
      "進行中": [],
      "即將到期": [],
      "已逾期": [],
      "已完成": []
    };

    for (var k = 0; k < 專案列表.length; k++) {
      var 專案 = 專案列表[k];
      var 狀態 = String(專案["狀態"] || "");
      var 顯示文字 = 專案["專案名稱"] + "\n(" + 專案["負責人"] + ")";

      if (狀態.indexOf("已完成") >= 0) {
        看板分類["已完成"].push(顯示文字);
      } else if (狀態.indexOf("已逾期") >= 0) {
        看板分類["已逾期"].push(顯示文字);
      } else if (狀態.indexOf("即將到期") >= 0) {
        看板分類["即將到期"].push(顯示文字);
      } else {
        看板分類["進行中"].push(顯示文字);
      }
    }

    // 3. 建立或清除「任務看板」工作表
    var 看板表 = ss.getSheetByName("任務看板");
    if (看板表) {
      看板表.clear();
    } else {
      看板表 = ss.insertSheet("任務看板");
    }

    // 4. 設定標題區塊
    看板表.getRange("A1").setValue("📋 專案任務看板");
    看板表.getRange("A1").setFontSize(16).setFontWeight("bold");
    看板表.getRange("A2").setValue("更新時間：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm"));

    // 5. 設定 Kanban 欄位標題
    var 欄位順序 = ["進行中", "即將到期", "已逾期", "已完成"];
    var 欄位標題 = [["🟢 進行中", "🟡 即將到期", "🔴 已逾期", "✅ 已完成"]];
    看板表.getRange(4, 1, 1, 4).setValues(欄位標題)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setFontSize(12)
      .setBackground("#eceff1");

    // 6. 找出最多任務的欄位，決定需要寫入幾列
    var 最大列數 = 0;
    for (var c = 0; c < 欄位順序.length; c++) {
      if (看板分類[欄位順序[c]].length > 最大列數) {
        最大列數 = 看板分類[欄位順序[c]].length;
      }
    }

    // 7. 填入看板資料
    if (最大列數 > 0) {
      var 輸出資料 = [];
      var 輸出背景 = [];
      // 定義各欄位的背景顏色
      var 背景對應 = ["#ffffff", "#fff3e0", "#ffebee", "#e8f5e9"];

      for (var r = 0; r < 最大列數; r++) {
        var 這一列 = [];
        var 這一列顏色 = [];
        for (var c = 0; c < 欄位順序.length; c++) {
          var 類別 = 欄位順序[c];
          var 內容 = 看板分類[類別][r] || "";
          這一列.push(內容);
          
          if (內容 !== "") {
            這一列顏色.push(背景對應[c]);
          } else {
            這一列顏色.push("#ffffff");
          }
        }
        輸出資料.push(這一列);
        輸出背景.push(這一列顏色);
      }

      var 範圍 = 看板表.getRange(5, 1, 最大列數, 4);
      範圍.setValues(輸出資料);
      範圍.setBackgrounds(輸出背景);
      範圍.setVerticalAlignment("middle");
      範圍.setHorizontalAlignment("center");
      範圍.setWrap(true);
      
      // 稍微加高列高以顯示換行內容
      for (var rr = 5; rr < 5 + 最大列數; rr++) {
        看板表.setRowHeight(rr, 45);
      }
    }

    // 8. 調整欄寬
    for (var col = 1; col <= 4; col++) {
      看板表.setColumnWidth(col, 180);
    }

    Logger.log("✅ 任務看板已生成！");
    SpreadsheetApp.getUi().alert("✅ 任務看板已生成！請查看「任務看板」工作表。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 智慧派工
 * 說明：自動依員工負擔與技能匹配分配未指派的任務，最少負擔與技能相符者優先。
 */
function 智慧派工() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 人員表 = ss.getSheetByName("專案人員");
    var 專案表 = ss.getSheetByName("專案追蹤");
    
    if (!人員表 || !專案表) {
      SpreadsheetApp.getUi().alert("❌ 找不到「專案人員」或「專案追蹤」工作表");
      return;
    }

    // 1. 讀取員工清單與職稱（作為技能依據）
    var 人員資料 = 人員表.getDataRange().getValues();
    var 員工名單 = [];
    var 員工職稱 = {};
    for (var i = 1; i < 人員資料.length; i++) {
      var 姓名 = 人員資料[i][0];
      員工名單.push(姓名); // A 欄：姓名
      員工職稱[姓名] = 人員資料[i][2]; // C 欄：職稱
    }

    if (員工名單.length === 0) {
      SpreadsheetApp.getUi().alert("❌ 沒有可用的員工資料");
      return;
    }

    // 2. 讀取專案資料並計算負擔
    var 專案資料 = 專案表.getDataRange().getValues();
    var 標題 = 專案資料[0];
    
    // 初始化負擔統計
    var 員工負擔 = {};
    員工名單.forEach(function(姓名) {
      員工負擔[姓名] = 0;
    });

    // 找出「負責人」、「進度(%)」與「專案名稱」的欄位索引
    var 負責人索引 = 標題.indexOf("負責人");
    var 進度索引 = 標題.indexOf("進度(%)");
    var 專案名稱索引 = 標題.indexOf("專案名稱");

    if (負責人索引 === -1 || 進度索引 === -1 || 專案名稱索引 === -1) {
      SpreadsheetApp.getUi().alert("❌ 工作表欄位格式不符");
      return;
    }

    // 計算每個人進行中的專案數
    for (var j = 1; j < 專案資料.length; j++) {
      var 負責人 = 專案資料[j][負責人索引];
      var 進度 = 專案資料[j][進度索引];
      
      // 如果有負責人且專案未完成，負擔 + 1
      if (負責人 && 進度 < 100 && 員工負擔.hasOwnProperty(負責人)) {
        員工負擔[負責人]++;
      }
    }

    // 定義技能匹配規則：職稱 -> 專案名稱關鍵字
    var 資格規則 = {
      "前端工程師": ["官網", "網頁", "介面"],
      "後端工程師": ["API", "資料庫", "系統"],
      "UI設計師": ["設計", "介面", "視覺"],
      "全端工程師": ["系統", "API", "完整", "改版"],
      "專案經理": ["管理", "協調", "規劃"],
      "UX研究員": ["用戶", "研究", "訪談"],
      "QA工程師": ["測試", "品質", "驗證"],
      "產品經理": ["產品", "規劃", "客戶"],
      "自動化測試": ["自動化", "測試"],
      "iOS工程師": ["App", "iOS", "手機"]
    };

    // 3. 尋找未指派的專案並進行派工
    var 有更新 = false;
    var 派工紀錄 = [];

    for (var k = 1; k < 專案資料.length; k++) {
      var 當前負責人 = 專案資料[k][負責人索引];
      var 專案名稱 = 專案資料[k][專案名稱索引];

      // 如果沒有負責人，進行智慧派工
      if (!當前負責人 || 當前負責人.toString().trim() === "") {
        
        // 3.1 找出符合技能的員工
        var 符合條件的員工 = [];
        for (var 員工 in 員工負擔) {
          var 職稱 = 員工職稱[員工];
          var 關鍵字 = 資格規則[職稱] || [];
          var 符合技能 = false;
          
          for (var r = 0; r < 關鍵字.length; r++) {
            if (專案名稱.indexOf(關鍵字[r]) >= 0) {
              符合技能 = true;
              break;
            }
          }
          
          if (符合技能) {
            符合條件的員工.push(員工);
          }
        }

        // 3.2 決定候選名單：優先用符合技能的，若無則用所有人（備援）
        var 候選名單 = 符合條件的員工.length > 0 ? 符合條件的員工 : Object.keys(員工負擔);
        
        // 3.3 從候選人中選出負擔最少的人
        var 最少負擔者 = null;
        var 最少負擔量 = Infinity;

        for (var m = 0; m < 候選名單.length; m++) {
          var 候選人 = 候選名單[m];
          if (員工負擔[候選人] < 最少負擔量) {
            最少負擔量 = 員工負擔[候選人];
            最少負擔者 = 候選人;
          }
        }

        if (最少負擔者) {
          // 填入負責人
          專案表.getRange(k + 1, 負責人索引 + 1).setValue(最少負擔者);
          // 更新負擔
          員工負擔[最少負擔者]++;
          有更新 = true;
          
          var 匹配類型 = 符合條件的員工.length > 0 ? " (技能匹配)" : " (負擔優先)";
          派工紀錄.push("「" + 專案名稱 + "」 ➡️ " + 最少負擔者 + 匹配類型);
        }
      }
    }

    if (有更新) {
      // 更新狀態以防新指派的專案狀態未計算
      定時更新專案狀態();
      SpreadsheetApp.getUi().alert("✅ 智慧派工完成！\n\n派工結果：\n" + 派工紀錄.join("\n"));
    } else {
      SpreadsheetApp.getUi().alert("ℹ️ 目前沒有未指派的專案。");
    }

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 發生錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化範例資料
// ============================================================

function 初始化專案資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 專案人員 ---
  var 人員表 = ss.getSheetByName("專案人員");
  if (!人員表) 人員表 = ss.insertSheet("專案人員"); else 人員表.clear();

  var 人員標題 = [["姓名", "部門", "職稱", "月薪", "Email"]];
  var 人員資料 = [
    ["王小明", "開發部", "前端工程師", 48000, "wang@company.com"],
    ["李小華", "開發部", "後端工程師", 52000, "lee@company.com"],
    ["張美玲", "設計部", "UI設計師", 45000, "chang@company.com"],
    ["陳大文", "開發部", "全端工程師", 58000, "chen@company.com"],
    ["林小芬", "企劃部", "專案經理", 55000, "lin@company.com"],
    ["黃志偉", "設計部", "UX研究員", 46000, "huang@company.com"],
    ["劉家豪", "測試部", "QA工程師", 42000, "liu@company.com"],
    ["吳雅琪", "企劃部", "產品經理", 60000, "wu@company.com"],
    ["周建國", "測試部", "自動化測試", 44000, "chou@company.com"],
    ["許文馨", "開發部", "iOS工程師", 56000, "hsu@company.com"]
  ];

  人員表.getRange(1, 1, 1, 5).setValues(人員標題);
  人員表.getRange(2, 1, 人員資料.length, 5).setValues(人員資料);
  人員表.getRange("A1:E1").setBackground("#673ab7").setFontColor("#fff").setFontWeight("bold");
  人員表.getRange("D2:D11").setNumberFormat("#,##0");
  人員表.setFrozenRows(1);
  for (var i = 1; i <= 5; i++) 人員表.autoResizeColumn(i);

  // --- 專案追蹤 ---
  var 專案表 = ss.getSheetByName("專案追蹤");
  if (!專案表) 專案表 = ss.insertSheet("專案追蹤"); else 專案表.clear();

  var 專案標題 = [["專案名稱", "負責人", "優先級", "開始日期", "截止日期", "進度(%)", "狀態"]];
  var 專案資料 = [
    ["官網改版", "林小芬", "高", new Date(2026, 3, 1), new Date(2026, 4, 15), 80, ""],
    ["App 2.0", "吳雅琪", "高", new Date(2026, 3, 10), new Date(2026, 5, 30), 35, ""],
    ["報表系統", "陳大文", "中", new Date(2026, 4, 1), new Date(2026, 4, 20), 100, ""],
    ["客戶管理", "王小明", "中", new Date(2026, 2, 15), new Date(2026, 4, 10), 60, ""],
    ["API 串接", "李小華", "高", new Date(2026, 4, 5), new Date(2026, 4, 25), 20, ""],
    ["設計系統", "張美玲", "低", new Date(2026, 3, 20), new Date(2026, 6, 1), 45, ""],
    ["自動化測試", "劉家豪", "中", new Date(2026, 4, 1), new Date(2026, 5, 15), 15, ""],
    ["用戶研究", "黃志偉", "低", new Date(2026, 3, 1), new Date(2026, 3, 30), 100, ""]
  ];

  專案表.getRange(1, 1, 1, 7).setValues(專案標題);
  專案表.getRange(2, 1, 專案資料.length, 7).setValues(專案資料);
  專案表.getRange("A1:G1").setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");
  專案表.getRange("D2:E9").setNumberFormat("yyyy/mm/dd");
  專案表.setFrozenRows(1);
  for (var j = 1; j <= 7; j++) 專案表.autoResizeColumn(j);

  SpreadsheetApp.getUi().alert("✅ 專案資料已建立！（專案人員 + 專案追蹤）");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 5 工具")
    .addItem("📦 初始化專案資料", "初始化專案資料")
    .addItem("📋 生成部門清單", "生成部門清單")
    .addItem("📊 更新專案狀態", "定時更新專案狀態")
    .addItem("📈 專案進度摘要", "生成專案摘要")
    .addItem("💰 生成薪資報表", "生成薪資報表")
    .addItem("📌 建立任務看板", "生成任務看板")
    .addItem("🤖 智慧派工", "智慧派工")
    .addItem("⏰ 檢查逾期任務", "彈窗提醒逾期任務")
    .addSeparator()
    .addItem("🔬 函數進階示範", "函數進階示範")
    .addItem("🗂️ 資料結構化示範", "資料結構化示範")
    .addItem("⏰ 設定定期更新", "設定定期更新觸發器")
    .addToUi();
}
