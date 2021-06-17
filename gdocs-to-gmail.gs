function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Convert')
      .addItem('To Gmail', 'convertToGmail')
      .addToUi();
}

function convertToGmail() {
  var ui = DocumentApp.getUi();
  ui.alert('convertToGmail called.');
}