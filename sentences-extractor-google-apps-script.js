/**
 * Sentences Extractor - Google Apps Script
 *
 * A Google Docs plugins that is capable of extracting
 * sentences in an essay to a new Google Docs document.
 *
 * @author Krerkkiat Chusap
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * Documentation is taken from the example code.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Extract sentences', 'onClickExtractSentences')
      .addItem('Extract and shuffles', 'onClickShuffledSentences')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * Documentation is taken from the example code.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Extarct sentenes, and write to new document.
 *
 */
function onClickExtractSentences() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var sentences = extractSentences(body);

  // New document.
  var new_doc = DocumentApp.create(doc.getName() + ' - Extracted Sentences');
  writeSentencesToNewDocument(new_doc, sentences);
  
  // Prompt the link.
  var ui = DocumentApp.getUi();
  var response = ui.alert('Link to a new document', 'This is a link to your extracted sentences: \n' + new_doc.getUrl(), ui.ButtonSet.OK);
}

/**
 * Extarct sentenes, shuffle them, and write to new document.
 *
 */
function onClickShuffledSentences() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var sentences = extractSentences(body);

  shuffleSentence(sentences);

  // New document.
  var new_doc = DocumentApp.create(doc.getName() + ' - Extracted Sentences');
  writeSentencesToNewDocument(new_doc, sentences);
  
  // Prompt the link.
  var ui = DocumentApp.getUi();
  var response = ui.alert('Link to a new document', 'This is a link to your extracted sentences: \n' + new_doc.getUrl(), ui.ButtonSet.OK);
}

/**
 * Shuffle extracted sentences.
 *
 * Code is heavily taken from https://bost.ocks.org/mike/shuffle/.
 *
 * @param {Array.<string>} sentences The array of sentences to be shuffle.
 * @return {Array.<string>} The array of shuffled sentences.
 */
function shuffleSentence(sentences) {
  var sentencesSize = sentences.length;
  var temp;
  var idx;

  // While there remain elements to shuffle.
  while (sentencesSize != 0) {
    // Pick a remaining element.
    idx = Math.floor(Math.random() * sentencesSize--);

    // And swap it with the current element.
    temp = sentences[sentencesSize];
    sentences[sentencesSize] = sentences[idx];
    sentences[idx] = temp;
  }
}

/**
 * Extracte sentences from document body.
 *
 * It will ignore first four pagragraphs in the document. In general essay, these four
 * lines usually contain author information, and the essay information.
 *
 * @param {Body} documentBody The body of the document that sentences will be extracted from.
 * @return {Array.<string>} The array of extracted sentences.
 */
function extractSentences(documentBody) {
  var paragraphContent = '';
  var sentences = [];
  var sentencesInParagraph;
  
  // Get all the paragraphs
  var paragraphs = documentBody.getParagraphs();
  
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

    sentencesInParagraph = paragraphContent.split('. ');
    for (var j = 0; j < sentencesInParagraph.length; j++) {
      sentences.push(sentencesInParagraph[j]);
    }
  }

  return sentences;
}

/**
 * Append sentences to the document.
 *
 * @param {Document} doc The document that sentences will be appended to.
 * @param {Array.<string>} sentences The array of sentences.
 */
function writeSentencesToNewDocument(doc, sentences) {
  var body = doc.getBody();

  for (var i = 0; i < sentences.length; i++) {
    body.appendParagraph(sentences[i] + '.');
    body.appendParagraph('');
  }

  doc.saveAndClose();
}