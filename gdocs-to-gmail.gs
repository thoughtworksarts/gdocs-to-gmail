var state = {
  body: null,
  testMode: false,
  testDocId: '1cuJMBUi9vrV3hbbwaccHatglD4EkfLPA9G1jYF9qzTw',
  isParsingActivated: false,
  currentElement: {
    obj: null,
    text: null
  },
  currentList: {
    isInProgress: false,
    type: '',
    items: []
  },
  outputLines: [],
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

function testApp() {
  state.testMode = true;
  convertToGmail();
}

function loadDocument() {
  if(state.testMode) {
    state.body = DocumentApp.openById(state.testDocId).getBody();
  } else {
    state.body = DocumentApp.getActiveDocument().getBody();
  }
}

function convertToGmail() {
  loadDocument();
  var numElements = state.body.getNumChildren();
  for(var i = 0; i < numElements; i++) {
    assignCurrentElement(state.body.getChild(i));
    deactivateParsingOnRangeEnd();
    parseWhileActive();
    activateParsingOnRangeBegin();
  }

  showResult();
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

function parseWhileActive() {
  if(state.isParsingActive) {
    parse(state.currentElement.obj);
  }
}

function parse(element) {
  switch(element.getType()) {
    case DocumentApp.ElementType.PARAGRAPH:
      closeListIfNeeded();
      parseParagraph(element.asParagraph());
      break;
    case DocumentApp.ElementType.LIST_ITEM:
      openListIfNeeded()
      parseListItem(element.asListItem());
      break;
    default:
      closeListIfNeeded();
      break;
  }
}

function parseParagraph(paragraph) {
  var numChildren = paragraph.getNumChildren();
  if(numChildren === 0) return '<br>';

  for(var i = 0; i < numChildren; i++) {
    var element = paragraph.getChild(i);
    switch(element.getType()) {
      case DocumentApp.ElementType.INLINE_IMAGE:
        parseInlineImage(element.asInlineImage());
        break;
      case DocumentApp.ElementType.TEXT:
        var heading = paragraph.getHeading();
        var text = element.asText();
        heading === DocumentApp.ParagraphHeading.NORMAL ? parseText(text) : parseHeading(text);
        break;
    }
  }
}

function parseHeading(textElement) {
  var str = textElement.getText();
  state.outputLines.push('<font size="4"><b>' + str + '</b></font>');
}

function parseText(textElement) {
  var html = '';
  var str = textElement.getText();
  var indices = textElement.getTextAttributeIndices();

  for (var i = 0; i < indices.length; i++) {
    var attributeStartIndex = indices[i];
    var attributeEndIndex = (i + 1 < indices.length) ? indices[i + 1] : str.length;
    var attribute = textElement.getAttributes(indices[i]);
    var substring = str.substring(attributeStartIndex, attributeEndIndex);

    html += attribute.BOLD ? '<b>' : '';
    html += attribute.ITALIC ? '<i>' : '';
    html += attribute.LINK_URL ? '<a href="' + attribute.LINK_URL + '">' : '';
    html += substring;
    html += attribute.LINK_URL ? '</a>' : '';
    html += attribute.BOLD ? '</b>' : '';
    html += attribute.ITALIC ? '</i>' : '';
  }
  state.outputLines.push(html);
}

function parseInlineImage(inlineImage) {
  state.outputLines.push('INLINE IMAGE');
}

function parseListItem(listItem) {
  appendCurrentOutputLine('<li>LIST_ITEM</li>\n');
}

function openListIfNeeded() {
  if(!state.currentList.isInProgress) {
    var listItem = state.currentElement.obj.asListItem();
    state.currentList.isInProgress = true;
    state.currentList.type = listItem.getGlyphType() === DocumentApp.GlyphType.NUMBER ? 'ol' : 'ul';
    state.outputLines.push('<' + state.currentList.type + '>\n');
  }
}

function closeListIfNeeded() {
  if(state.currentList.isInProgress) {
    var listItemType = state.currentList.type;
    state.currentList.isInProgress = false;
    state.currentList.type = '';
    appendCurrentOutputLine('</' + listItemType + '>');
  }
}

function appendCurrentOutputLine(str) {
  state.outputLines[state.outputLines.length - 1] += str;
}

function wrapWithHeaderFooter(bodyHtml) {
  var headerHtml = HtmlService.createHtmlOutputFromFile('header').getContent();
  var footerHtml = HtmlService.createHtmlOutputFromFile('footer').getContent();
  return headerHtml + '\n' + bodyHtml + '\n' + footerHtml;
}

function showResult() {
  var resultStr = '<div>' + state.outputLines.join('</div>\n<div>') + '</div>'
  resultStr = wrapWithHeaderFooter(resultStr);
  if(state.testMode) {
    Logger.log(resultStr);
  } else {
    DocumentApp.getUi().alert(resultStr);
  }
}