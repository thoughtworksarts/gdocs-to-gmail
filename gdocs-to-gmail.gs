var state = {
  body: DocumentApp.getActiveDocument().getBody(),
  isParsingActivated: false,
  currentElement: {
    obj: null,
    text: null
  },
  returnHtml: '',
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
  var numElements = state.body.getNumChildren();
  for(var i = 0; i < numElements; i++) {
    assignCurrentElement(state.body.getChild(i));
    deactivateParsingOnRangeEnd();
    parseWhenActive();
    activateParsingOnRangeBegin();
  }

  DocumentApp.getUi().alert(wrap(state.returnHtml));
}

function assignCurrentElement(element) {
    state.currentElement.obj = element;
    state.currentElement.text = element.asText().getText();
}

function parseWhenActive() {
  if(state.isParsingActive) {
    state.returnHtml += state.currentElement.obj.getType() + ' ' + state.currentElement.text + '\n';
  }
}

function activateParsingOnRangeBegin() {
  if(state.currentElement.text === state.rangeMarkers.begin) {
    state.isParsingActive = true;
  }
}

function deactivateParsingOnRangeEnd() {
  if(state.currentElement.text === state.rangeMarkers.end) {
    state.isParsingActive = false;
  }
}

function wrap(bodyHtml) {
  var headerHtml = HtmlService.createHtmlOutputFromFile('header').getContent();
  var footerHtml = HtmlService.createHtmlOutputFromFile('footer').getContent();
  return headerHtml + '\n' + bodyHtml + '\n' + footerHtml;
}