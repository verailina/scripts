/* This script is automatically executed when on a corresponding google document (to which it has been attached).
The script adds a new report entry at the end of the document and scrolls up to its last line. An enty format:

Mars, 20
AM:
  * <to fill>
PM:
  * <to fill>
A new entry always corresponds to the last buisiness day before the current date. (Holydays are not processed).
*/

/* Check if the year is leap */
function leapYear(year)
{
  return ((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0);
}

/* Get last work date */
function getLastWorkDate(currentDate) {
  months = ["Jan", "Fev", "Mars", "Avr", "Mai", "Juin", "Juil", "Aout", "Sept", "Oct", "Nov", "Dec"];
  nMonthDays = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (currentDate == undefined){
    var currentDate = new Date();
  }
  yesterday = currentDate.getDate();
  month = currentDate.getMonth();
  year = currentDate.getFullYear();
  
  do {
    yesterday -= 1;
    if (yesterday == 0) {
      month -= 1;
      if (month == -1){
        month = 11;
        year -= 1;
      }
      if (leapYear(year)) {
        nMonthDays[1] += 1;
      }
      yesterday = nMonthDays[month];
    }
    prevDay = new Date(year, month, yesterday);
  } while (prevDay.getDay() == 0 || prevDay.getDay() == 6);
  return {day: yesterday, month: months[month], year: year, newWeek: (prevDay.getDay() == 1)};
}

/*Append a new paragraph with the given text and make it bold if necessary. */
function addNewParagraph(textStr, body, isBold=false) {
  var paragraph = body.appendParagraph(textStr);
  text = paragraph.editAsText();
  text.setBold(isBold);
  return paragraph
}

/* Add new list item <to fill>. */
function addNewListItem(body) {
   body.appendListItem("<to fill>").setNestingLevel(0).setGlyphType(DocumentApp.GlyphType.BULLET);
}

/* Add new report entry for the last work date. */
function addNewEntry(date, body) {
  var paragraph = addNewParagraph(`${date.month}, ${date.day}`, body, date.newWeek);
  addNewParagraph("AM:", body);
  addNewListItem(body);
  addNewParagraph("PM:", body);
  addNewListItem(body);
  var lastParagraph = addNewParagraph("", body);
  return lastParagraph;
}

/* Add a new report entry and scroll at the last page when the document is oppened. */
function onOpen(e) {
  yesterday = getLastWorkDate();
  doc = DocumentApp.getActiveDocument();
  body = doc.getBody();
  var lastParagraph = addNewEntry(yesterday, body);

  var position = doc.newPosition(lastParagraph, 0);
  doc.setCursor(position);

}

/* Execute on installation.  */
function onInstall(e) {
  onOpen(e);
}
