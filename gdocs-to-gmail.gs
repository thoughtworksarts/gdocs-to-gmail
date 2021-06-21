var state = {
  body: null,
  testMode: false,
  testDocId: '1cuJMBUi9vrV3hbbwaccHatglD4EkfLPA9G1jYF9qzTw',
  isProcessingActive: false,
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
    deactivateProcessingOnRangeEnd();
    processWhileActive();
    activateProcessingOnRangeBegin();
  }

  showResult();
}

function assignCurrentElement(element) {
    state.currentElement.obj = element;
    state.currentElement.text = element.asText().getText();
}

function activateProcessingOnRangeBegin() {
  if(state.currentElement.text === state.rangeMarkers.begin) {
    state.isProcessingActive = true;
  }
}

function deactivateProcessingOnRangeEnd() {
  if(state.currentElement.text === state.rangeMarkers.end) {
    state.isProcessingActive = false;
  }
}

function processWhileActive() {
  if(state.isProcessingActive) {
    process(state.currentElement.obj);
  }
}

function process(element) {
  switch(element.getType()) {
    case DocumentApp.ElementType.PARAGRAPH:
      closeListIfNeeded();
      processParagraph(element.asParagraph());
      break;
    case DocumentApp.ElementType.LIST_ITEM:
      openListIfNeeded()
      processListItem(element.asListItem());
      break;
    default:
      closeListIfNeeded();
      break;
  }
}

function processParagraph(paragraph) {
  var numChildren = paragraph.getNumChildren();
  if(numChildren === 0) state.outputLines.push('<br>');

  for(var i = 0; i < numChildren; i++) {
    var element = paragraph.getChild(i);
    switch(element.getType()) {
      case DocumentApp.ElementType.INLINE_IMAGE:
        processInlineImage(element.asInlineImage());
        break;
      case DocumentApp.ElementType.TEXT:
        var heading = paragraph.getHeading();
        var text = element.asText();
        heading === DocumentApp.ParagraphHeading.NORMAL ? processText(text) : processHeading(text);
        break;
    }
  }
}

function processHeading(textElement) {
  var str = textElement.getText();
  state.outputLines.push('<font size="4"><b>' + str + '</b></font>');
}

function processText(textElement) {
  state.outputLines.push(parseText(textElement));
}

function processInlineImage(inlineImage) {
  state.outputLines.push('INLINE IMAGE');
}

function processListItem(listItem) {
  var html = '';
  var numChildren = listItem.getNumChildren();
  if(numChildren === 0) html = '&nbsp;';

  for(var i = 0; i < numChildren; i++) {
    var element = listItem.getChild(i);
    switch(element.getType()) {
      case DocumentApp.ElementType.TEXT:
        html += parseText(element.asText());
        break;
    }
  }
  appendCurrentOutputLine('<li>' + html + '</li>\n');
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
    html += attribute.UNDERLINE && !attribute.LINK_URL ? '<u>' : '';
    html += attribute.LINK_URL ? '<a href="' + attribute.LINK_URL + '">' : '';
    html += substring;
    html += attribute.LINK_URL ? '</a>' : '';
    html += attribute.UNDERLINE && !attribute.LINK_URL ? '</u>' : '';
    html += attribute.ITALIC ? '</i>' : '';
    html += attribute.BOLD ? '</b>' : '';
  }

  return html.replace('\r', '<br>');
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