/**
 * Sentences Extractor - Google Apps Script
 *
 * A Google Docs plugins that is capable of extracting
 * sentences in essay to a new Google Docs document.
 *
 * @author Krerkkiat Chusap
 */

function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Extract sentences', 'extractSentences')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function extractSentences() {
  // Current document.
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  // New document.
  var new_doc = DocumentApp.create(doc.getName() + ' - Extracted Sentences');
  var new_body = new_doc.getBody();
  
  var pattern = /^\t.*/gi;
  var paragraphContent = '';
  
  //var allSentences = [];
  var sentences;
  
  // Get all the paragraphs
  var paragraphs = body.getParagraphs();
  
  for (var i = 4; i < paragraphs.length; i++) {
    paragraphContent = paragraphs[i].getText().trim();
    
    // Stop if we reach the reference page.
    if (paragraphContent.indexOf('Reference') == 0) {
      break;
    }
    
    // Ignore the blank paragraphs.
    if (paragraphContent == '') {
      continue;
    }
    
    sentences = paragraphContent.split('. ');
    for (var j = 0; j < sentences.length; j++) {
      new_body.appendParagraph(sentences[j] + '.');
      new_body.appendParagraph('');
    }
  }
  
  new_doc.saveAndClose();
  
  // Prompt the link.
  var ui = DocumentApp.getUi();
  var response = ui.alert('Link to a new document', 'This is a link to your extracted sentences: \n' + new_doc.getUrl(), ui.ButtonSet.OK);
  
}