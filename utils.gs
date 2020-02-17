// date(momentオブジェクト)が休日かどうか判定
// https://qiita.com/jz4o/items/d4e978f9085129155ca6
function isHoliday(date) {
  
  Logger.log(date.toDate());
  
  // 土日なら候補に入れる
  var weekInt = date.day();
  var result = false;
  if (weekInt <= 0 || 6 <= weekInt) {
    result = true;
  }
  
  // 祝日なら候補に入れる
  var calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  var calendar = CalendarApp.getCalendarById(calendarId);
  var todayEvents = calendar.getEventsForDay(date.toDate());
  Logger.log(todayEvents);
  if (todayEvents.length > 0) {
    result = true;
  }
  
  // 本番なら弾く
  var CNstageCalendarId = PropertiesService.getScriptProperties("STAGE_CALENDAR_ID");
  var CNstageCalendar = CalendarApp.getCalendarById(CNstageCalendarId);
  Logger.log(CNstageCalendar.getName());
  var todayStage = CNstageCalendar.getEventsForDay(date.toDate());
  if (todayStage.length > 0) {
    result = false;
  }
  
  return result;
}
