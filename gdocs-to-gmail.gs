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

function parseWhenActive() {
  if(state.isParsingActive) {
    state.returnHtml += parse(state.currentElement.obj);
  }
}

function parse(element) {
  switch(element.getType()) {
    case DocumentApp.ElementType.PARAGRAPH:
      return parseParagraph(element.asParagraph());
    case DocumentApp.ElementType.LIST_ITEM:
      return parseListItem(element.asListItem());
    default:
      return '';
  }
}

function parseParagraph(paragraph) {
  var numChildren = paragraph.getNumChildren();
  if(numChildren === 0) return '<div><br></div>\n';

  var returnHtml = '';

  for(var i = 0; i < numChildren; i++) {
    var element = paragraph.getChild(i);
    switch(element.getType()) {
      case DocumentApp.ElementType.INLINE_IMAGE:
        returnHtml += parseInlineImage(element.asInlineImage());
        break;
      case DocumentApp.ElementType.TEXT:
        returnHtml += parseText(element.asText());
        break;
    }
  }

  return returnHtml;
}

function parseText(text) {
  var returnHtml = '<div>';
  var str = text.getText();
  var indices = text.getTextAttributeIndices();

  for (var i = 0; i < indices.length; i++) {
    var attributeStartIndex = indices[i];
    var attributeEndIndex = (i + 1 < indices.length) ? indices[i + 1] : str.length;
    var attribute = text.getAttributes(indices[i]);
    var substring = str.substring(attributeStartIndex, attributeEndIndex);

    returnHtml += attribute.BOLD ? '<b>' : '';
    returnHtml += attribute.ITALIC ? '<i>' : '';
    returnHtml += substring;
    returnHtml += attribute.BOLD ? '</b>' : '';
    returnHtml += attribute.ITALIC ? '</i>' : '';
  }
  return returnHtml + '</div>\n';
}

function parseInlineImage(inlineImage) {
  return 'INLINE IMAGE\n';
}

function parseListItem(listItem) {
  return 'LIST ITEM\n';
}

function wrap(bodyHtml) {
  var headerHtml = HtmlService.createHtmlOutputFromFile('header').getContent();
  var footerHtml = HtmlService.createHtmlOutputFromFile('footer').getContent();
  return headerHtml + '\n' + bodyHtml + '\n' + footerHtml;
}