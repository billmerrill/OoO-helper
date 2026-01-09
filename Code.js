/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current document.
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


function getSettings() {
  return NSPropertyHelper.getAll();
}

function saveSettings(settings) {
  const saveSettings = {'eventTitle': settings['eventTitle'],
                         'p1Name': settings['p1Name'],
                         'p2Name': settings['p2Name']}
  NSPropertyHelper.setAll(saveSettings);
}
