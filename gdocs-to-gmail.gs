var state = {
  body: DocumentApp.getActiveDocument().getBody(),
  rangeMarkers: {
    begin: '===CUSTOM_EMAIL_CONTENTS_BEGIN===',
    end: '===CUSTOM_EMAIL_CONTENTS_END==='
  }
}

function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Convert')
      .addItem('To Gmail', 'convertToGmail')
      .addToUi();
}

function convertToGmail() {
  var str = '';
  var isActive = false;
  var numElements = state.body.getNumChildren();
  for(var i = 0; i < numElements; i++) {
    var element = state.body.getChild(i);
    var elementText = element.asText().getText();
    if(elementText == state.rangeMarkers.end) isActive = false;
    if(isActive) str += element.getType() + ' ' + elementText + '\n';
    if(elementText == state.rangeMarkers.begin) isActive = true;
  }

  DocumentApp.getUi().alert('convertToGmail called.\n\n' + wrap(str));
}

function wrap(bodyHtml) {
  var headerHtml = HtmlService.createHtmlOutputFromFile('header').getContent();
  var footerHtml = HtmlService.createHtmlOutputFromFile('footer').getContent();
  return headerHtml + '\n' + bodyHtml + '\n' + footerHtml;
}