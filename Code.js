function estimateSpeechTime(text) {
  // Increase words per minute for Bahasa Indonesia
  var wordsPerMinute = 250;

  // Reduce the impact of syllables
  var secondsPerSyllable = 0.075;

  // Slightly reduce pause times
  var pauseForComma = 0.1;
  var pauseForPeriod = 0.2;

  var words = text.split(/\s+/).length;

  // Adjust syllable counting for Bahasa Indonesia
  var syllables = (text.match(/[aiueoAIUEO]/g) || []).length;

  var commas = (text.match(/,/g) || []).length;
  var periods = (text.match(/\./g) || []).length;

  var time =
    (words / wordsPerMinute) * 60 +
    syllables * secondsPerSyllable +
    commas * pauseForComma +
    periods * pauseForPeriod;

  return time;
}

function getNormalText(body) {
  var normalText = '';
  var numChildren = body.getNumChildren();
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = child.asParagraph();
      if (paragraph.getHeading() == DocumentApp.ParagraphHeading.NORMAL) {
        normalText += paragraph.getText() + ' ';
      }
    }
  }
  return normalText.trim();
}

function getJakartaTime() {
  var date = new Date();
  var jakartaTime = Utilities.formatDate(
    date,
    'GMT+7',
    'dd MMMM yyyy HH:mm:ss'
  );
  return jakartaTime;
}

function appendOutputToDocument(doc, output) {
  var body = doc.getBody();

  // Find the previous output paragraph
  var paragraphs = body.getParagraphs();
  var outputParagraph = null;
  for (var i = paragraphs.length - 1; i >= 0; i--) {
    if (paragraphs[i].getText().startsWith('Estimasi Durasi Bicara:')) {
      outputParagraph = paragraphs[i];
      break;
    }
  }

  // Append new output with current date and time
  var jakartaTime = getJakartaTime();
  var fullOutput = 'Estimasi Durasi Bicara: ' + jakartaTime + '\n' + output;

  if (outputParagraph) {
    // Update existing paragraph
    outputParagraph.setText(fullOutput);
    outputParagraph.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  } else {
    // Append new paragraph
    var newParagraph = body.appendParagraph(fullOutput);
    newParagraph.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  }
}

function estimateFullSpeechTime() {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();

    var normalText = getNormalText(body);
    var duration = estimateSpeechTime(normalText);
    var minutes = Math.floor(duration / 60);
    var seconds = Math.floor(duration % 60);

    var output =
      'Durasi (hanya teks normal): ' + minutes + ' menit ' + seconds + ' detik';

    DocumentApp.getUi().alert(output);
    appendOutputToDocument(doc, output);
  } catch (e) {
    Logger.log('An error occurred: ' + e.toString());
    DocumentApp.getUi().alert('An error occurred: ' + e.toString());
  }
}

function estimateSelectedSpeechTime() {
  try {
    var selection = DocumentApp.getActiveDocument().getSelection();
    if (selection) {
      var selectedElements = selection.getSelectedElements();
      var selectedText = '';
      for (var i = 0; i < selectedElements.length; i++) {
        var element = selectedElements[i].getElement();
        if (element.editAsText) {
          var text = element.editAsText();
          var textString = text.getText();

          // Check if the element is a paragraph and has Normal style
          if (
            element.getType() === DocumentApp.ElementType.PARAGRAPH &&
            element.asParagraph().getHeading() ===
              DocumentApp.ParagraphHeading.NORMAL
          ) {
            selectedText += textString + ' ';
          } else if (element.getType() === DocumentApp.ElementType.TEXT) {
            // If it's a text element, check its parent paragraph's style
            var parentParagraph = element.getParent();
            if (
              parentParagraph.getType() === DocumentApp.ElementType.PARAGRAPH &&
              parentParagraph.asParagraph().getHeading() ===
                DocumentApp.ParagraphHeading.NORMAL
            ) {
              selectedText += textString + ' ';
            }
          }
        }
      }
      selectedText = selectedText.trim();

      if (selectedText) {
        var duration = estimateSpeechTime(selectedText);
        var minutes = Math.floor(duration / 60);
        var seconds = Math.floor(duration % 60);

        var output =
          'Durasi teks terpilih (hanya gaya Normal): ' +
          minutes +
          ' menit ' +
          seconds +
          ' detik';
        DocumentApp.getUi().alert(output);
      } else {
        DocumentApp.getUi().alert(
          'Tidak ada teks dengan gaya Normal yang dipilih.'
        );
      }
    } else {
      DocumentApp.getUi().alert('Silakan pilih teks terlebih dahulu.');
    }
  } catch (e) {
    Logger.log('An error occurred: ' + e.toString());
    DocumentApp.getUi().alert('An error occurred: ' + e.toString());
  }
}

function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Speech Time')
    .addItem('Estimasi Durasi Bicara Total', 'estimateFullSpeechTime')
    .addItem(
      'Estimasi Durasi Bicara Teks Terpilih',
      'estimateSelectedSpeechTime'
    )
    .addToUi();
}
