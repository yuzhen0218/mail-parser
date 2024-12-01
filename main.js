function processEmails() {
  var startDate = new Date("2023-03-01"); // 替換成起始日期
  var endDate = new Date("2023-03-03"); // 替換成結束日期
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("提交");
  var searchString = "subject:學習歷程資料庫 學分異常檢查通知";
  var start = 0; // 從搜尋結果的第一封郵件開始
  var max = 500; // 返回最多 500 封符合條件的郵件
  var threads = GmailApp.search(searchString, start, max);
  var foundEmailsCount = 0; // 初始化郵件計數器

  // 建立一個陣列來追蹤已經存在的成績名冊 FID 和重修重讀名冊 FID 的組合
  var existingFIDPairs = [];

  // 取得已存在的資料
  var existingData = sheet.getRange("C:D").getValues();
  existingData.forEach(function(row) {
    var gradeFID = row[0].toString().trim();
    var reexamFID = row[1].toString().trim();
    var key = gradeFID + "_" + reexamFID; // 使用成績名冊 FID 和重修重讀名冊 FID 的組合作為 key
    existingFIDPairs.push(key);
  });

  threads.sort((a, b) => a.getLastMessageDate() - b.getLastMessageDate()).forEach(function(thread) {
    var messages = thread.getMessages();
    messages.forEach(function(message) {
      var receivedTime = message.getDate();

      if (receivedTime.getTime() >= startDate.getTime() && receivedTime.getTime() <= endDate.getTime()) {
        var content = message.getPlainBody();
        var formattedReceivedTime = Utilities.formatDate(receivedTime, "GMT+8", "yyyy-MM-dd HH:mm:ss");

        var schoolCodeMatch = content.match(/學校：.*?-(.*?)-/);
        var schoolNameMatch = content.match(/-.*?-(.*?)，/);
        var abnormalCountMatch = content.match(/共發現(\d+)筆/);
        var rosterFIDMatch = content.match(/FID：(\d+)/);
        var gradeFIDMatch = content.match(/學期成績FID：(\d+)/);
        
        var gradeFID; // 宣告 gradeFID 變數，用於存放學期成績FID的結果
        var reexamFID; // 宣告 reexamFID 變數，用於存放重修重讀名冊 FID 或進修部成績名冊 FID 的結果

        if (!schoolCodeMatch || !schoolNameMatch || !abnormalCountMatch || !rosterFIDMatch) {
          return; // 如果 match 失敗，跳過這封信件，繼續處理下一封信件
        }

        var schoolCode = schoolCodeMatch[1];
        var schoolName = schoolNameMatch[1].trim();
        var abnormalCount = abnormalCountMatch[1];
        var rosterFID = rosterFIDMatch[1];
        
        // 檢查學期成績FID是否存在
        if (gradeFIDMatch) {
          gradeFID = gradeFIDMatch[1];
        } else {
          gradeFID = ""; // 若學期成績FID不存在，將 gradeFID 設為空白
        }

        // 在填入 Google Sheets 時，使用單引號包裹學校代碼
        var schoolCodeString = "'" + schoolCode.toString().padStart(6, "0");

        // 檢查 schoolCodeString 是否為 "000000"
        if (schoolCodeString === "'000000") {
          return; // 如果學校代碼是 "000000"，跳過該筆資料，繼續處理下一封信件
        }

        // 檢查重修重讀名冊 FID 或進修部成績名冊 FID 是否存在
        if (rosterFIDMatch) {
          reexamFID = rosterFIDMatch[1];
        } else {
          reexamFID = "";
        }

        // 檢查是否已經存在相同的成績名冊 FID 或重修重讀名冊 FID 的組合
        var key = gradeFID + "_" + reexamFID; // 使用成績名冊 FID 和重修重讀名冊 FID 的組合作為 key
        if (existingFIDPairs.includes(key)) {
          return; // 如果已經存在相同 FID 的資料，跳過這封信件，繼續處理下一封信件
        } else {
          existingFIDPairs.push(key); // 將新的 key 加入陣列，表示已經處理過這個 FID 組合
        }

        // 調換 rosterFID 和 gradeFID 的順序
        var rowData = [schoolCodeString, schoolName, gradeFID, reexamFID, formattedReceivedTime, abnormalCount];
        sheet.appendRow(rowData); // 將 rowData 新增到 "提交" 工作表的最後一行
        foundEmailsCount++;
      }
    });
  });

  Logger.log("找到符合條件的信件數量：" + foundEmailsCount);
}
