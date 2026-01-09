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

