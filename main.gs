/**
 * 開啟 Spreadsheet 的時候新增 Menu
 */
function onOpen(e) {
  //選單名稱
  var m = SpreadsheetApp.getUi().createMenu('Google Calendar Functions');
  //選單項目
  m.addItem('get calendar ids', 'getids');
  m.addItem('delete selected calendar events', 'delSelCalEvent');
  m.addToUi();
}
/**
 * 取得所有日曆
 */
function getids() {
  //取得目前所在的工作表
  var sht = SpreadsheetApp.getActiveSheet();
  //清除工作表內容
  sht.clear();
  //取得所有日曆
  var ary = CalendarApp.getAllOwnedCalendars();
  if (ary.length === 0) {
    sht.getRange(1, 1).setValue("找不到已存在的日曆");
    return;  
  }
  //設定標題
  sht.getRange(1, 1, 1, 2).setValues([["ID", "名稱"]]);
  //顯示所有日曆
  for (var i=0; i<ary.length; i++){
    sht.getRange(i + 2, 1, 1, 2).setValues([[ary[i].getId(), ary[i].getName()]]);
  }
  //設定欲刪除的日曆輸入框
  sht.getRange(ary.length + 2, 1).setValue("請輸入欲刪除的日曆 ID:");
  sht.getRange(ary.length + 2 + 1, 1).setValue("請輸入起始日期(ex:1900/01/01):");
  sht.getRange(ary.length + 2 + 2, 1).setValue("請輸入結束日期(ex:2078/12/31):");
  sht.getRange(ary.length + 2 + 3, 1).setValue("請輸入欲刪除事件名稱：");
    
}
/**
 * 刪除指定日曆下的所有事件
 */
function delSelCalEvent(){
  //取得目前所在的工作表
  var sht = SpreadsheetApp.getActiveSheet();
  var lastrow = sht.getLastRow();
  var idRow = 0;
  for (var j=1; j<lastrow; j++){
    if (sht.getRange(j, 1).getValue() == "請輸入欲刪除的日曆 ID:"){
      idRow = j;
    }
  }
  
  if (idRow > 0){
    //取得欲刪除的日曆 ID
    var calid = sht.getRange(idRow, 2).getValue();
    //取得該日曆物件
    var cal = CalendarApp.getCalendarById(calid);
    var beginDate = new Date(sht.getRange(idRow + 1, 2).getValue());
    var endDate = new Date(sht.getRange(idRow + 2, 2).getValue());
    var eventTitle = sht.getRange(idRow + 3, 2).getValue();
    if (cal != null){
      var ary = cal.getEvents(beginDate, endDate);
      for (var i=0; i<ary.length; i++){
        sht.getRange(idRow + i + 1, 1).setValue(ary[i].getTitle());
        if(ary[i].getTitle() == eventTitle){
          ary[i].deleteEvent();
          sht.getRange(idRow + i + 1, 2).setValue("已刪除");
        }
      }
    }
  }
}

