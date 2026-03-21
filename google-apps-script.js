// ONE桌遊營運管理儀表板 - Google Apps Script API
// 部署為網頁應用程式，讓儀表板可以讀取資料

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = e.parameter.sheet || '目標追蹤';
  
  try {
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        error: '找不到工作表：' + sheetName
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({
        values: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 轉換為物件陣列
    const headers = data[0];
    const rows = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] !== undefined ? row[index] : '';
      });
      return obj;
    });
    
    return ContentService.createTextOutput(JSON.stringify({
      values: rows,
      count: rows.length
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 測試函數
function testAPI() {
  const result = doGet({ parameter: { sheet: '目標追蹤' } });
  Logger.log(result.getContent());
}
