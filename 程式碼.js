function calculateSafeStockByColorAndGroup() {
  // 1. 取得來源工作表「2」
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('8');
  if (!sheet) {
    Logger.log('指定的工作表不存在！');
    return;
  }

  // 2. 取得資料範圍、背景顏色及數值資料
  var dataRange = sheet.getDataRange();
  var bgColors = dataRange.getBackgrounds();
  var data = dataRange.getValues();
  var targetColor = "#ead1dc"; // 淺洋紅色3的顏色代碼

  // 3. 以廠商（來源工作表 P 欄，索引15）為 key 分組，儲存符合條件的資料
  var vendorGroups = {};

  // 從第2行開始遍歷（跳過標題）
  for (var i = 1; i < bgColors.length; i++) {
    // 如果該行任何一格背景色符合目標顏色
    if (bgColors[i].some(function(color) { return color.toLowerCase() === targetColor; })) {
      var drugName = data[i][1];        // 藥名：B 欄（索引1）
      var currentBalance = parseFloat(data[i][11]); // 目前結餘：L 欄（索引11）
      var safeStock = parseFloat(data[i][12]);     // 安全庫存：M 欄（索引12）
      var vendor = data[i][15];       // 廠商：P 欄（索引15）

      // 檢查 parseFloat 的結果是否為 NaN
      if (!isNaN(safeStock) && !isNaN(currentBalance)) {
        var doubleSafeStock = safeStock * 2;
        // 原始程式中僅用來標示結餘是否小於安全庫存（此處未加入輸出，如有需要可額外處理）
        var checkMark = (currentBalance < safeStock) ? "✔" : "";

        // 使用模板字串來簡化字串拼接
        var drugInfoString = `${drugName} (兩倍安全庫存: ${doubleSafeStock})`;

        // 將資料依廠商分組，若該廠商尚無群組則先建立
        if (!vendorGroups[vendor]) {
          vendorGroups[vendor] = [];
        }
        vendorGroups[vendor].push({
          entryString: drugInfoString,
          currentBalance: currentBalance,
          checkMark: checkMark
        });
      } else {
          // 處理 NaN 的情況，例如記錄錯誤
          Logger.log(`第 ${i + 1} 行的安全庫存或目前結餘資料有誤`);
      }
    }
  }

  // 4. 建立輸出結果陣列
  // 標題可依需求調整，這裡僅呈現【廠商】及合併後的【藥品資訊】
  var results = [["廠商", "藥品資訊"]];

  // 依每個廠商將同群組的藥品資訊合併到同一儲存格（以換行符分隔）
  for (var vendor in vendorGroups) {
    var combinedEntry = vendorGroups[vendor]
      .map(function(item) { return item.entryString; })
      .join("\n");
    results.push([vendor, combinedEntry]);
  }

  // 5. 輸出結果至工作表【結果】（若不存在則自動建立，並先清除舊資料）
  var resultSheet = ss.getSheetByName('結果') || ss.insertSheet('結果');
  resultSheet.clear();
  resultSheet.getRange(1, 1, results.length, results[0].length).setValues(results);

  Logger.log('已成功處理符合條件的行，依廠商分組並合併藥名+兩倍安全庫存資訊至儲存格！');
}