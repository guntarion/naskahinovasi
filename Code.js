function estimateSpeechTime(text) {
  // Average reading speed in words per minute
  var wordsPerMinute = 150;

  // Average time to say a syllable in seconds
  var secondsPerSyllable = 0.15;

  // Time for a pause in seconds
  var pauseForComma = 0.15;
  var pauseForPeriod = 0.3;

  // Count the number of words
  var words = text.split(/\s+/).length;

  // Count the number of syllables (this is a simple approximation)
  var syllables = (text.match(/[aeiouAEIOU]/g) || []).length;

  // Count the number of commas and periods
  var commas = (text.match(/,/g) || []).length;
  var periods = (text.match(/\./g) || []).length;

  // Calculate the total time in seconds
  var time =
    (words / wordsPerMinute) * 60 +
    syllables * secondsPerSyllable +
    commas * pauseForComma +
    periods * pauseForPeriod;

  return time;
}

function main() {
  try {
    var doc = DocumentApp.getActiveDocument();
    var text = doc.getBody().getText();

    var duration = estimateSpeechTime(text);
    var minutes = Math.floor(duration / 60);
    var seconds = Math.floor(duration % 60);

    var output = 'Durasi: ' + minutes + ' menit ' + seconds + ' detik';

    // Display the result in a dialog box
    DocumentApp.getUi().alert(output);

    // Optionally, you can also add the result to the end of the document
    doc.getBody().appendParagraph(output);
  } catch (e) {
    Logger.log('An error occurred: ' + e.toString());
    DocumentApp.getUi().alert('An error occurred: ' + e.toString());
  }
}

function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Speech Time')
    .addItem('Estimate Speech Time', 'main')
    .addToUi();
}
