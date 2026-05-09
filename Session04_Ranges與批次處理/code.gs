// ============================================================
// Session 4：Ranges 與批次處理
// 日期：115/05/09（六）13:30~16:30
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. Ranges（儲存格範圍）存取與修改
//   2. 執行限制（批次處理最佳化）
//   3. 陣列 (Array) 基礎
//   4. 實作：批次讀寫資料，自動匯入範例
// ============================================================

// ============================================================
// 第一部分：Range 進階操作
// ============================================================

/**
 * Range 取得方式大全
 * 說明：示範各種 getRange 的用法
 */
function Range取得方式() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 請先執行「初始化庫存資料」");
    return;
  }

  // 方式 1：A1 表示法
  var r1 = sheet.getRange("A1");
  Logger.log("A1 = " + r1.getValue());

  // 方式 2：A1 範圍表示法
  var r2 = sheet.getRange("A1:C3");
  Logger.log("A1:C3 = " + JSON.stringify(r2.getValues()));

  // 方式 3：數字表示法 getRange(列, 欄)
  var r3 = sheet.getRange(1, 1);  // 等同 A1
  Logger.log("(1,1) = " + r3.getValue());

  // 方式 4：數字範圍 getRange(起始列, 起始欄, 列數, 欄數)
  var r4 = sheet.getRange(2, 1, 5, 4);  // 從 A2 開始，5列4欄
  Logger.log("(2,1,5,4) = " + JSON.stringify(r4.getValues()));

  // 方式 5：整欄 / 整列
  var r5 = sheet.getRange("A:A");  // 整個 A 欄
  var r6 = sheet.getRange("2:2");  // 整個第 2 列
  Logger.log("第 2 列 = " + JSON.stringify(r6.getValues()));
}

/**
 * Range 修改操作
 * 說明：修改儲存格的值與屬性
 */
function Range修改操作() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
  if (!sheet) return;

  // --- 單一儲存格操作 ---
  var cell = sheet.getRange("A1");
  cell.setValue("商品名稱");
  cell.setBackground("#e8f5e9");
  cell.setFontWeight("bold");
  cell.setFontSize(12);
  cell.setHorizontalAlignment("center");
  cell.setNote("這是商品名稱欄位");  // 加入附註

  // --- 範圍操作 ---
  var range = sheet.getRange("B2:B11");
  range.setNumberFormat("#,##0");       // 數字格式：千分位
  range.setHorizontalAlignment("right"); // 靠右

  // --- 使用公式 ---
  // F2 = D2 * E2（庫存量 × 單價）
  sheet.getRange("F2").setFormula("=D2*E2");

  // 批次設定公式（F2:F11）
  for (var i = 2; i <= 11; i++) {
    sheet.getRange("F" + i).setFormula("=D" + i + "*E" + i);
  }

  Logger.log("✅ Range 修改完成");
}

// ============================================================
// 第二部分：陣列 (Array) 基礎
// ============================================================

/**
 * 陣列基礎操作示範
 */
function 陣列基礎() {
  // --- 一維陣列 ---
  var 水果 = ["蘋果", "香蕉", "橘子", "葡萄", "芒果"];

  Logger.log("陣列長度：" + 水果.length);          // 5
  Logger.log("第一個元素：" + 水果[0]);             // 蘋果
  Logger.log("最後一個：" + 水果[水果.length - 1]); // 芒果

  // 新增元素
  水果.push("西瓜");
  Logger.log("push 後：" + 水果);            // 末尾加入

  // 刪除最後一個
  var 被移除 = 水果.pop();
  Logger.log("pop 移除：" + 被移除);          // 西瓜

  // 搜尋元素
  var 索引 = 水果.indexOf("橘子");
  Logger.log("橘子的索引：" + 索引);           // 2

  // --- 二維陣列（模擬試算表資料）---
  var 表格 = [
    ["姓名", "年齡", "部門"],     // 標題列
    ["王小明", 25, "業務部"],
    ["李小華", 30, "行銷部"],
    ["張美玲", 28, "人資部"]
  ];

  Logger.log("第 2 列第 1 欄：" + 表格[1][0]); // 王小明
  Logger.log("第 3 列第 3 欄：" + 表格[2][2]); // 行銷部

  // 遍歷二維陣列
  for (var i = 1; i < 表格.length; i++) {
    Logger.log(表格[i][0] + "，" + 表格[i][1] + " 歲，" + 表格[i][2]);
  }
}

// ============================================================
// 第三部分：批次處理最佳化
// ============================================================

/**
 * ❌ 不良示範：逐格讀寫（效能極差）
 * 說明：每次 getValue/setValue 都是一次 API 呼叫，
 *       大量操作時會非常慢且容易超時
 */
function 不良示範_逐格讀寫() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
  if (!sheet) return;

  var 開始 = new Date().getTime();

  // ❌ 這樣做很慢！每次迴圈都會呼叫 API
  for (var i = 2; i <= 11; i++) {
    var 商品 = sheet.getRange(i, 1).getValue();  // API 呼叫 ×1
    var 單價 = sheet.getRange(i, 2).getValue();  // API 呼叫 ×1
    var 庫存 = sheet.getRange(i, 4).getValue();  // API 呼叫 ×1
    var 總值 = 單價 * 庫存;
    sheet.getRange(i, 6).setValue(總值);          // API 呼叫 ×1
    // 每列 4 次 API 呼叫，10 列 = 40 次！
  }

  var 結束 = new Date().getTime();
  Logger.log("❌ 逐格讀寫耗時：" + (結束 - 開始) + " 毫秒");
}

/**
 * ✅ 正確示範：批次讀寫（效能最佳）
 * 說明：一次讀取全部、處理完再一次寫回
 */
function 正確示範_批次讀寫() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
  if (!sheet) return;

  var 開始 = new Date().getTime();

  // ✅ 一次讀取所有資料（1 次 API 呼叫）
  var 所有資料 = sheet.getDataRange().getValues();

  // 在記憶體中處理（0 次 API 呼叫）
  var 結果 = [];
  for (var i = 1; i < 所有資料.length; i++) {
    var 單價 = 所有資料[i][1];   // B 欄：單價
    var 庫存 = 所有資料[i][3];   // D 欄：庫存量
    var 總值 = 單價 * 庫存;
    結果.push([總值]);
  }

  // ✅ 一次寫入所有結果（1 次 API 呼叫）
  sheet.getRange(2, 6, 結果.length, 1).setValues(結果);

  var 結束 = new Date().getTime();
  Logger.log("✅ 批次讀寫耗時：" + (結束 - 開始) + " 毫秒");
  // 通常比逐格讀寫快 10 倍以上！
}

// ============================================================
// 第四部分：批次匯入外部資料
// ============================================================

/**
 * 自動匯入範例資料（模擬外部資料匯入）
 * 說明：示範如何將大量資料批次寫入試算表
 */
function 批次匯入資料() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 表名 = "匯入資料_" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmmss");
    var sheet = ss.insertSheet(表名);

    // 模擬產生 50 筆訂單資料
    var 標題 = [["訂單編號", "客戶名稱", "商品", "數量", "單價", "金額", "日期", "狀態"]];
    var 資料 = [];

    var 客戶列表 = ["台北公司", "新竹企業", "台中行銷", "高雄科技", "花蓮文創"];
    var 商品列表 = ["筆記型電腦", "印表機", "投影機", "平板電腦", "耳機", "滑鼠", "鍵盤"];
    var 狀態列表 = ["已出貨", "處理中", "已取消", "已完成"];

    for (var i = 0; i < 50; i++) {
      var 訂單編號 = "ORD-" + String(i + 1).padStart(4, "0");
      var 客戶 = 客戶列表[Math.floor(Math.random() * 客戶列表.length)];
      var 商品 = 商品列表[Math.floor(Math.random() * 商品列表.length)];
      var 數量 = Math.floor(Math.random() * 10) + 1;
      var 單價 = Math.floor(Math.random() * 50000) + 500;
      var 金額 = 數量 * 單價;
      var 日期 = new Date(2026, Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1);
      var 狀態 = 狀態列表[Math.floor(Math.random() * 狀態列表.length)];

      資料.push([訂單編號, 客戶, 商品, 數量, 單價, 金額, 日期, 狀態]);
    }

    // 批次寫入標題
    sheet.getRange(1, 1, 1, 8).setValues(標題);

    // 批次寫入資料（一次寫入 50 筆！）
    sheet.getRange(2, 1, 資料.length, 8).setValues(資料);

    // 格式化
    var 標題範圍 = sheet.getRange("A1:H1");
    標題範圍.setBackground("#ea4335");
    標題範圍.setFontColor("#ffffff");
    標題範圍.setFontWeight("bold");
    標題範圍.setHorizontalAlignment("center");

    // 數字格式
    sheet.getRange("E2:F51").setNumberFormat("#,##0");
    sheet.getRange("G2:G51").setNumberFormat("yyyy/mm/dd");

    sheet.setFrozenRows(1);

    // --- 新增：依照「金額」(第 6 欄) 由大到小排序 ---
    var 資料範圍 = sheet.getRange(2, 1, 資料.length, 8);
    資料範圍.sort({column: 6, ascending: false});

    for (var j = 1; j <= 8; j++) {
      sheet.autoResizeColumn(j);
    }

    Logger.log("✅ 已匯入並排序完成，共 " + 資料.length + " 筆訂單。");
    SpreadsheetApp.getUi().alert("✅ 已匯入 " + 資料.length + " 筆訂單！\n工作表：" + 表名);

  } catch (錯誤) {
    Logger.log("❌ 匯入錯誤：" + 錯誤.message);
  }
}

/**
 * 批次更新庫存（讀取→計算→寫回）
 * 說明：一次讀取所有庫存，計算庫存總值後批次寫回
 */
function 批次更新庫存() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("庫存資料");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 找不到「庫存資料」");
      return;
    }

    // Step 1：一次讀取全部資料
    var 資料 = sheet.getDataRange().getValues();

    // Step 2：在記憶體中計算
    var 庫存總值列 = [];
    var 庫存狀態列 = [];
    var 總庫存值 = 0;
    var 預警清單 = []; // 用來收集庫存不足的商品

    for (var i = 1; i < 資料.length; i++) {
      var 單價 = 資料[i][1];     // B: 單價
      var 安全庫存 = 資料[i][2]; // C: 安全庫存量
      var 目前庫存 = 資料[i][3]; // D: 目前庫存
      var 庫存總值 = 單價 * 目前庫存;

      庫存總值列.push([庫存總值]);

      // 判斷庫存狀態
      var 狀態;
      if (目前庫存 <= 0) {
        狀態 = "🔴 缺貨";
        預警清單.push("❌ 【缺貨】" + 資料[i][0] + " (目前：" + 目前庫存 + ")");
      } else if (目前庫存 < 安全庫存) {
        狀態 = "🟡 低庫存";
        預警清單.push("⚠️ 【低庫存】" + 資料[i][0] + " (目前：" + 目前庫存 + "，安全水位：" + 安全庫存 + ")");
      } else {
        狀態 = "🟢 正常";
      }
      庫存狀態列.push([狀態]);

      總庫存值 += 庫存總值;
    }

    // --- 新增：發送 Email 預警 ---
    if (預警清單.length > 0) {
      var 收件人 = Session.getActiveUser().getEmail(); // 發送給目前操作者
      var 主旨 = "【庫存預警】系統檢測到有商品庫存不足";
      var 內容 = "您好，系統偵測到以下商品庫存已低於安全水位：\n\n" + 
                 預警清單.join("\n") + 
                 "\n\n請及時進行採購。\n\n(此為系統自動發送，請勿直接回覆)";
      
      MailApp.sendEmail(收件人, 主旨, 內容);
      Logger.log("📧 已發送預警信至：" + 收件人);
    }

    // Step 3：一次寫回所有結果
    var 資料筆數 = 資料.length - 1;
    sheet.getRange(2, 6, 資料筆數, 1).setValues(庫存總值列);  // F 欄：庫存總值
    sheet.getRange(2, 7, 資料筆數, 1).setValues(庫存狀態列);  // G 欄：狀態

    // 數字格式
    sheet.getRange(2, 6, 資料筆數, 1).setNumberFormat("#,##0");

    // 在最後寫入總計
    var 總計列 = 資料.length + 1;
    sheet.getRange(總計列, 5).setValue("總庫存值：");
    sheet.getRange(總計列, 5).setFontWeight("bold");
    sheet.getRange(總計列, 6).setValue(總庫存值);
    sheet.getRange(總計列, 6).setNumberFormat("#,##0").setFontWeight("bold").setBackground("#e8f5e9");

    Logger.log("✅ 庫存更新完成！總庫存值：NT$" + 總庫存值.toLocaleString());
    SpreadsheetApp.getUi().alert("✅ 庫存更新完成！\n總庫存值：NT$ " + 總庫存值.toLocaleString());

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化範例資料
// ============================================================

/**
 * 建立「庫存資料」工作表
 */
function 初始化庫存資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("庫存資料");

  if (!sheet) {
    sheet = ss.insertSheet("庫存資料");
  } else {
    sheet.clear();
  }

  var 標題 = [["商品名稱", "單價", "安全庫存量", "目前庫存", "單位", "庫存總值", "狀態"]];
  var 資料 = [
    ["A4 影印紙", 150, 50, 120, "包", "", ""],
    ["原子筆", 15, 100, 45, "支", "", ""],
    ["資料夾", 25, 80, 200, "個", "", ""],
    ["釘書針", 35, 30, 8, "盒", "", ""],
    ["便條紙", 20, 60, 0, "本", "", ""],
    ["膠帶", 30, 40, 55, "捲", "", ""],
    ["白板筆", 45, 20, 15, "支", "", ""],
    ["計算機", 350, 10, 12, "台", "", ""],
    ["印表機碳粉", 1200, 5, 3, "個", "", ""],
    ["滑鼠墊", 80, 15, 25, "個", "", ""]
  ];

  sheet.getRange(1, 1, 1, 7).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 7).setValues(資料);

  var 標題範圍 = sheet.getRange("A1:G1");
  標題範圍.setBackground("#ff6d01");
  標題範圍.setFontColor("#ffffff");
  標題範圍.setFontWeight("bold");
  標題範圍.setHorizontalAlignment("center");

  sheet.getRange("B2:B11").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  for (var i = 1; i <= 7; i++) sheet.autoResizeColumn(i);

  Logger.log("✅ 庫存資料已建立！");
  SpreadsheetApp.getUi().alert("✅ 庫存資料已建立！\n請執行「批次更新庫存」計算結果。");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 4 工具")
    .addItem("👥 初始化員工資料", "初始化員工資料")
    .addItem("📦 初始化庫存資料", "初始化庫存資料")
    .addItem("📊 批次更新庫存", "批次更新庫存")
    .addItem("📥 批次匯入訂單資料", "批次匯入資料")
    .addItem("➕ 批次去重匯入新客戶", "批次去重匯入")
    .addSeparator()
    .addItem("📝 Range 取得方式", "Range取得方式")
    .addItem("🎨 Range 修改操作", "Range修改操作")
    .addItem("📚 陣列基礎", "陣列基礎")
    .addSeparator()
    .addItem("📸 建立手動快照", "建立庫存快照")
    .addItem("⏰ 開啟每日自動快照", "設定每日快照觸發器")
    .addSeparator()
    .addItem("🔍 批次驗證資料內容", "批次驗證資料")
    .addItem("🧹 批次清理與格式化", "批次清理資料")
    .addSeparator()
    .addItem("❌ 不良示範：逐格讀寫", "不良示範_逐格讀寫")
    .addItem("✅ 正確示範：批次讀寫", "正確示範_批次讀寫")
    .addToUi();
}

/**
 * 批次驗證資料：Email、電話、必填、重複
 * 說明：一次讀取所有資料，在記憶體中驗證後批次寫回，並標示錯誤儲存格
 */
function 批次驗證資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 嘗試取得「員工資料」或「待驗證資料」
  var sheet = ss.getSheetByName("員工資料") || ss.getSheetByName("待驗證資料");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到驗證對象（員工資料），請先建立資料。");
    return;
  }

  // 1. 一次讀取所有資料
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert("⚠️ 沒有資料可以驗證。");
    return;
  }

  var rows = values.slice(1); // 排除標題

  // 欄位索引預設 (依 Session 3 員工資料結構)
  var idx姓名 = 0;
  var idx電話 = 5;
  var idxEmail = 6;
  var idx結果 = 7; // H 欄

  var 結果列表 = [];
  var 背景顏色 = [];
  var email計數器 = {};

  // 正規表達式
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var phoneRegex = /^09\d{2}-?\d{3}-?\d{3}$/;

  // 2. 批次驗證處理
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var 錯誤訊息 = [];
    // 初始化該列背景色（null 表示維持原樣）
    var rowColor = new Array(idx結果 + 1).fill(null); 

    var 姓名 = String(row[idx姓名]).trim();
    var 電話 = String(row[idx電話]).trim();
    var Email = String(row[idxEmail]).trim();

    // A. 必填檢查
    if (!姓名) 錯誤訊息.push("姓名必填");
    if (!Email) 錯誤訊息.push("Email必填");

    // B. Email 格式檢查
    if (Email && !emailRegex.test(Email)) {
      錯誤訊息.push("Email 格式錯誤");
      rowColor[idxEmail] = "#ffebee"; // 標註 Email 欄位為粉紅
    }

    // C. 電話格式檢查
    if (電話 && !phoneRegex.test(電話)) {
      錯誤訊息.push("電話格式錯誤");
      rowColor[idx電話] = "#ffebee";
    }

    // D. 重複偵測 (Email)
    if (Email && emailRegex.test(Email)) {
      if (email計數器[Email]) {
        錯誤訊息.push("Email 重複");
        rowColor[idxEmail] = "#fff9c4"; // 標註重複為淡黃
      } else {
        email計數器[Email] = true;
      }
    }

    // 彙整驗證狀態
    if (錯誤訊息.length > 0) {
      結果列表.push(["❌ " + 錯誤訊息.join(", ")]);
      rowColor[idx結果] = "#ffcdd2"; // 結果欄標為紅色
    } else {
      結果列表.push(["✅ 通過"]);
      rowColor[idx結果] = "#c8e6c9"; // 結果欄標為綠色
    }
    背景顏色.push(rowColor);
  }

  // 3. 批次寫回結果 (包含格式與值)
  sheet.getRange(1, idx結果 + 1).setValue("驗證結果").setFontWeight("bold").setBackground("#eeeeee");
  
  var resultRange = sheet.getRange(2, idx結果 + 1, 結果列表.length, 1);
  resultRange.setValues(結果列表);
  resultRange.setHorizontalAlignment("left");

  // 一次性設定所有儲存格背景色 (非常高效)
  sheet.getRange(2, 1, 背景顏色.length, idx結果 + 1).setBackgrounds(背景顏色);
  
  sheet.autoResizeColumn(idx結果 + 1);

  Logger.log("✅ 批次驗證完成，共檢查 " + rows.length + " 筆資料。");
  SpreadsheetApp.getUi().alert("✅ 批次驗證完成！\n\n紅色：格式錯誤或必填缺失\n黃色：重複資料\n綠色：驗證通過");
}

/**
 * 批次清理資料：去空白、統一小寫、電話格式化
 */
function 批次清理資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("員工資料");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到「員工資料」工作表。");
    return;
  }

  // 1. 一次讀取所有資料
  var range = sheet.getDataRange();
  var values = range.getValues();
  var header = values[0];
  var rows = values.slice(1);

  // 欄位索引 (電話:5, Email:6)
  var idx電話 = 5;
  var idxEmail = 6;

  // 2. 批次處理資料
  var 清理後資料 = rows.map(function(row) {
    return row.map(function(cell, index) {
      // 排除日期物件，只針對字串處理
      if (cell instanceof Date) return cell;
      
      var strCell = String(cell).trim(); // 所有欄位去空白

      // A. Email 統一轉小寫
      if (index === idxEmail) {
        strCell = strCell.toLowerCase();
      }

      // B. 電話格式化 (0912345678 -> 0912-345-678)
      if (index === idx電話) {
        // 先去掉所有非數字字元
        var numbers = strCell.replace(/\D/g, "");
        if (numbers.length === 10 && numbers.startsWith("09")) {
          strCell = numbers.substring(0, 4) + "-" + numbers.substring(4, 7) + "-" + numbers.substring(7, 10);
        }
      }

      return strCell;
    });
  });

  // 3. 一次寫回資料 (不含標題)
  if (清理後資料.length > 0) {
    // 寫入純文字資料
    sheet.getRange(2, 1, 清理後資料.length, header.length).setValues(清理後資料);
    
    // 清理完成後，將背景顏色重設（清除驗證時產生的標色）
    sheet.getRange(2, 1, 清理後資料.length, header.length + 1).setBackground(null);
    
    SpreadsheetApp.getUi().alert("✅ 批次清理完成！\n1. 已去除多餘空格\n2. Email 已統一小寫\n3. 電話已標準化為 0000-000-000\n4. 已清除錯誤標色背景");
  }
}

/**
 * 批次去重匯入：檢查重複 Email 並只匯入新客戶
 */
function 批次去重匯入() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("員工資料");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到「員工資料」工作表。");
    return;
  }

  // 1. 取得現有所有 Email (用於比對去重)
  var existingData = sheet.getDataRange().getValues();
  var existingEmails = new Set();
  
  // 從第 2 列開始讀取 Email (索引 6)
  for (var i = 1; i < existingData.length; i++) {
    var email = String(existingData[i][6]).trim().toLowerCase();
    if (email) existingEmails.add(email);
  }

  // 2. 準備待匯入的新資料 (模擬外部來源)
  var 待匯入資料 = [
    ["趙六", "業務部", "專員", "2024/5/10", 35000, "0988-111-222", "zhao@example.com"], // 新人
    ["王小明", "業務部", "專員", "2023/3/15", 38000, "0912-345-678", "wang@example.com"], // 重複者 (初始化資料中已有)
    ["孫七", "行銷部", "專員", "2024/5/12", 36000, "0977-333-444", "sun@example.com"]   // 新人
  ];

  // 3. 過濾掉已存在的資料
  var 真正匯入清單 = 待匯入資料.filter(function(row) {
    var newEmail = String(row[6]).trim().toLowerCase();
    return !existingEmails.has(newEmail); // 只有不存在於現有集合中的才保留
  });

  // 4. 批次追加到表格底部
  if (真正匯入清單.length > 0) {
    var 最後一列 = sheet.getLastRow();
    sheet.getRange(最後一列 + 1, 1, 真正匯入清單.length, 7).setValues(真正匯入清單);
    
    SpreadsheetApp.getUi().alert("✅ 匯入完成！\n本次偵測到 " + (待匯入資料.length - 真正匯入清單.length) + " 筆重複，\n成功匯入 " + 真正匯入清單.length + " 筆新客戶。");
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ 沒有新資料需要匯入（所有資料皆已存在）。");
  }
}

/**
 * 建立「員工資料」測試樣本（含故意設計的錯誤資料供驗證測試）
 */
function 初始化員工資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("員工資料") || ss.insertSheet("員工資料");
  sheet.clear();

  var 標題 = [["姓名", "部門", "職稱", "到職日", "月薪", "電話", "Email"]];
  var 資料 = [
    ["王小明", "業務部", "專員", "2023/3/15", 38000, "0912-345-678", "wang@example.com"],
    ["李小華", "行銷部", "主管", "2021/8/1", 52000, "0923-456-78", "lee@example.com"], // 電話少一碼
    ["張美玲", "人資部", "專員", "2022/11/20", 40000, "0934-567-890", "chang#example.com"], // Email 格式錯
    ["陳大文", "研發部", "工程師", "2024/1/10", 55000, "0945-678-901", "chen@example.com"],
    ["林小芬", "財務部", "會計", "2020/6/15", 42000, "0956-789-012", "lin@example.com"],
    ["黃志偉", "業務部", "主管", "2019/4/1", 58000, "0967-890-123", "wang@example.com"], // Email 與王小明重複
    ["", "研發部", "實習生", "2024/5/1", 28000, "0978-901-234", "test@test.com"] // 姓名缺失
  ];

  sheet.getRange(1, 1, 1, 7).setValues(標題);
  sheet.getRange(1, 1, 1, 7).setBackground("#4285f4").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange(2, 1, 資料.length, 7).setValues(資料);
  sheet.setFrozenRows(1);
  for (var i = 1; i <= 7; i++) sheet.autoResizeColumn(i);

  SpreadsheetApp.getUi().alert("✅ 員工測試資料已建立！\n\n內含故意設計的錯誤（重複、格式錯、缺失），\n現在請點選選單執行「🔍 批次驗證資料內容」。");
}

/**
 * 建立每日庫存快照
 * 說明：將「庫存資料」完整複製一份，並以日期命名
 */
function 建立庫存快照() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 來源表 = ss.getSheetByName("庫存資料");
  
  if (!來源表) {
    SpreadsheetApp.getUi().alert("❌ 找不到「庫存資料」工作表，無法建立快照。");
    return;
  }

  // 1. 產生日期字串 (例如：20260509)
  var 今天 = new Date();
  var 日期字串 = Utilities.formatDate(今天, "Asia/Taipei", "yyyyMMdd");
  var 快照名稱 = "庫存快照_" + 日期字串;

  // 2. 檢查是否已有同名的快照，若有則先刪除
  var 舊快照 = ss.getSheetByName(快照名稱);
  if (舊快照) {
    ss.deleteSheet(舊快照);
  }

  // 3. 執行複製
  var 新快照 = 來源表.copyTo(ss);
  新快照.setName(快照名稱);

  // 4. 將快照表移到最後面，避免干擾主要操作
  var 所有表 = ss.getSheets();
  ss.setActiveSheet(新快照);
  ss.moveActiveSheet(所有表.length);

  Logger.log("✅ 庫存快照已建立：「" + 快照名稱 + "」");
  
  // 嘗試彈出視窗（若在觸發器環境執行會自動忽略）
  try {
    SpreadsheetApp.getUi().alert("✅ 庫存快照已建立！\n名稱：" + 快照名稱);
  } catch (e) {}
}

/**
 * 設定每日凌晨自動建立快照的觸發器
 */
function 設定每日快照觸發器() {
  // 先刪除舊的，避免重複
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "建立庫存快照") ScriptApp.deleteTrigger(t);
  });

  // 建立新觸發器：每天凌晨 00:00 ~ 01:00 執行
  ScriptApp.newTrigger("建立庫存快照")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  SpreadsheetApp.getUi().alert("✅ 每日自動快照已設定！\n系統將於每天凌晨 00:00~01:00 自動執行備份。");
}
