/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current document.
 */

/**
 * UI Management
 */
var SIDEBAR_TITLE = 'OoO Settings';

function onOpen(e) {
  DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Settings', 'showSidebar')
// DEV ONLY      .addItem('Run Test', 'test_PropStore')  
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Agenda Magic
 */
function checkOrScheduleNext1on1() {
  const result = processSchedulingLogic();
  const ui = DocumentApp.getUi();
  if (result.success) {
    ui.alert(`Success! Scheduled for ${result.date}.`);
  } else {
    ui.alert(result.message);
  }
}


/**
 * Finds the child index of the first Heading 2 element with specific text.
 * @param {string} targetText The exact text to look for (e.g., "One on One Meetings").
 * @returns {number} The index of the element, or -1 if not found.
 */
function findHeadingIndex(targetText) {
  const body = DocumentApp.getActiveDocument().getBody();
  const numChildren = body.getNumChildren();
  
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    
    // Check if the element is a paragraph
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = child.asParagraph();
      
      // Check if it is Heading 2 and matches the text
      if (paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING1 && 
          paragraph.getText().trim() === targetText) {
        return i;
      }
    }
  }
  
  return -1; // Not found
}

function getInsertionPoint() {
  insertIndex = findHeadingIndex("One on One Meetings");
  if (insertIndex == -1) {
    insertIndex = 1
  }
  return insertIndex;
}

function processSchedulingLogic() {
  const settings = getSettings();
 
  if (('lastMeetingDate' in settings) && !isDateInPast(settings.lastMeetingDate)) {
    return { success: false, message: "A future meeting is already scheduled." };
  }
  
  const nextMeetingDate = findNextMeetingDate(settings.eventTitle); 
  if (!nextMeetingDate) {
    return { success: false, message: "Could not find calendar event: " + settings.eventTitle };
  }
 
  const insertPt = getInsertionPoint();
  const newDateString = insertNewMeeting(settings, insertPt, nextMeetingDate);
  NSPropertyHelper.set('lastMeetingDate', newDateString);
  return { success: true, date: newDateString };
}

function insertNewMeeting(settings, insertionPoint, nextMeetingDate) {
  const timeZone = Session.getScriptTimeZone();
  const dateString = Utilities.formatDate(nextMeetingDate, timeZone, 'yyyy-MM-dd');
  const body = DocumentApp.getActiveDocument().getBody();
  
  const h2 = (ip, text) => body.insertParagraph(ip, text).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const h3 = (ip, text) => body.insertParagraph(ip, text).setBold(true);
  const li = (ip, text) => body.insertListItem(ip, text).setBold(false).setGlyphType(DocumentApp.GlyphType.BULLET);

  let idx = insertionPoint + 1;
  h2(idx++, dateString);
  h3(idx++, `${values.p1Name}'s Topics`);
  li(idx++, '');
  h3(idx++, `${values.p2Name}'s Topics`);
  li(idx++, '');
  h3(idx++, "Notes");
  li(idx++, '');
  body.insertHorizontalRule(idx++);

  return dateString;
}

function isDateInPast(dateString) {
  if (!dateString) return true; 
  const storedDate = new Date(dateString); 
  const today = new Date();
  storedDate.setHours(0, 0, 0, 0);
  today.setHours(0, 0, 0, 0);
  return storedDate.getTime() < today.getTime();
}

function findNextMeetingDate(eventName) {
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(new Date(), new Date(Date.now() + 30 * 24 * 60 * 60 * 1000), { search: eventName });
  return (events.length > 0) ? events[0].getStartTime() : null;
}