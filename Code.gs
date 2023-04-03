/* ***********************************************************************************
 * 
 *  Bible linker for Google Docs v2
 * 
 *  A Google Documents Apps Script that searches for Bible verses and creates links to
 *  a selection of online Bible sources.
 *
 *  For more information, visit: https://github.com/majal/bible-linker-google-docs
 *
 *********************************************************************************** */

const BIBLE_DATA_DEFAULT='https://github.com/majal/bible-linker-google-docs/raw/linker-v2-commenter/bible-data/en_jw.json';

function main(bibleDataSelection, bibleVersion) {

  // Initialize Google Document
  var doc = DocumentApp.getActiveDocument();

  // Initialize access to Docs UI
  var ui = DocumentApp.getUi();

  // Note if document has active selection
  var docSelection = doc.getSelection();

  // Initialize user preferences on Bible data source to use
  const userProperties = PropertiesService.getUserProperties();

  // Fetch Bible data, then save data source to user preferences
  if ( ! bibleDataSelection ) {
    var bibleData = JSON.parse(UrlFetchApp.fetch(BIBLE_DATA_DEFAULT));
    userProperties.setProperty('bibleData', BIBLE_DATA_DEFAULT);
  } else {
    var bibleData = JSON.parse(UrlFetchApp.fetch(bibleDataSelection));
    userProperties.setProperty('bibleData', bibleDataSelection);
  };

  // Get single-chapter Bible books
  var bookSingleChapters = bibleData.bookSingleChapters;

  // Get search RegExes
  var searchRegexMultiChapters = bibleData.searchRegex.multiChapters;
  var searchRegexSingleChapters = bibleData.searchRegex.singleChapters;

  // Get error messages
  var errorMsgParserTitle = bibleData.strings.errorMessages.parserError.title;
  var errorMsgParserBefore = bibleData.strings.errorMessages.parserError.messageBefore;
  var errorMsgParserAfter = bibleData.strings.errorMessages.parserError.messageAfter;

  //Initialize variables to use in loops
  var bookNum, bookName, searchRegex;

  // Get book numbers, process each
  for (bookNum of Object.keys(bibleData.bookNames)) {

    // Get book names, process each
    for (bookName of Object.values(bibleData.bookNames[bookNum])) {
      
      // Modify RegEx format if it is a single-chapter book
      if ( bookSingleChapters.includes(bookNum) ) {
        searchRegex = bookName + ' ' + searchRegexSingleChapters;
      } else {
        searchRegex = bookName + ' ' + searchRegexMultiChapters;
      };
      
      ///////////////////////////////////////////////////////////////////////////
      // For each bookName, each selection or whole document, perform a search //
      ///////////////////////////////////////////////////////////////////////////

      // Initialize variables
      var searchElement, searchResult, searchResultTextSlice;

      // Search only selected text if present
      if (docSelection) {
        var rangeElements = docSelection.getRangeElements();

        // Search within each selection element and pass results to parser 
        for (let n=0; n < rangeElements.length; n++) {
          searchElement = rangeElements[n].getElement();
          searchResult = searchElement.findText(searchRegex);

          // Note if selection is only a part of the element (single line)
          let searchElementStart = rangeElements[n].getStartOffset();
          let searchElementEnd = rangeElements[n].getEndOffsetInclusive();
          
          // Send found matches to parser
          try {

            // If whole line is selected start and end will be -1
            // https://developers.google.com/apps-script/reference/document/range-element#getendoffsetinclusive
            if (searchElementStart == -1 && searchElementEnd == -1) {
              bibleParse(bibleData, bibleVersion, bookNum, bookName, searchResult, searchElement, searchRegex);
            } else {
              bibleParse(bibleData, bibleVersion, bookNum, bookName, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd);
            };
            
          } catch {

            // Show text where search error occurred
            searchResultTextSlice = searchResult.getElement().asText().getText().slice(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive() + 1);
            // ui.alert(errorMsgParserTitle, errorMsgParserBefore + searchResultTextSlice + errorMsgParserAfter, ui.ButtonSet.OK);

          };

        };

      // Perform search on whole document
      } else {

        searchElement = doc.getBody();
        searchResult = searchElement.findText(searchRegex);
        
        // Send found matches to parser
        try {

          bibleParse(bibleData, bibleVersion, bookNum, bookName, searchResult, searchElement, searchRegex);
        
        } catch {

          // Show text where search error occurred
          searchResultTextSlice = searchResult.getElement().asText().getText().slice(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive() + 1);
          // ui.alert(errorMsgParserTitle, errorMsgParserBefore + searchResultTextSlice + errorMsgParserAfter, ui.ButtonSet.OK);

        };

      }; // END: For each bookName, each selection or whole document, perform a search

    }; // END: Get book names, process each

  }; // END: Get book numbers, process each

}; // END: function main(bibleDataSelection)


function bibleParse(bibleData, bibleVersion, bookNum, bookName, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd) {

  // Pull default Bible data if nothing was provided
  if ( ! bibleData ) bibleData = JSON.parse(UrlFetchApp.fetch(BIBLE_DATA_DEFAULT));

  // Variable(s) and constant(s)
  var bookSingleChapters = bibleData.bookSingleChapters;
  var matchRegex = [ '[0-9]+:[0-9]+-[0-9]+:[0-9]+', '[0-9]+:[0-9]+, [0-9]+', '[0-9]+:[0-9]+-[0-9]+', '[0-9]+:[0-9]+' ];
  var spaceChars = '[ \u00a0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u200b\u2028\u2029\u3000]';

  // Cycle through each Bible reference found
  while (searchResult != null) {

    // Get reference start and end offsets
    let searchResultStart = searchResult.getStartOffset();
    let searchResultEnd = searchResult.getEndOffsetInclusive();

    // Isolate reference text
    let searchResultAstext = searchResult.getElement().asText();
    let searchResultText = searchResultAstext.getText();
    let searchResultTextSlice = searchResultText.slice(searchResultStart, searchResultEnd + 1);

    Logger.log('|' + searchResultTextSlice + '|');


    // Match searchResult with matchRegex-es, longest to shortest
    let matchResult = null;
    let i = 0;
    while ( ! matchResult && i < matchRegex.length ) {      
      matchResult = searchResultTextSlice.match(bookName + spaceChars + matchRegex[i]);
      i++;

      // Check if matchResult is comma separated
      if ( matchResult && matchResult[0].match(spaceChars + '[0-9]+:[0-9]+,' + spaceChars + '[0-9]+') ) {
        let vStart = matchResult[0].match(':[0-9]+')
        vStart = parseInt(vStart[0].replace(':', ''), 10);

        let vEnd = matchResult[0].match('[0-9]+$')
        vEnd = parseInt(vEnd[0].replace(':', ''), 10);

        // If comma-separated and non-consecutive, discard matchResult
        if ( vStart + 1 != vEnd ) matchResult = null;
      };

    };

    let chapterStart, verseStart, verseEnd, chapterEnd, url;

    chapterStart = matchResult[0].match('[0-9]+:');
    
    if ( chapterStart.length > 1 ) {
      chapterEnd = chapterStart[1].match('[0-9]+:');
      chapterEnd = chapterEnd[0].replace(':', '');
    };

    chapterStart = chapterStart[0].replace(':', '');

    verseStart = matchResult[0].match(':[0-9]+');
    verseEnd = matchResult[0].match('[0-9]+$');
    
    url = getUrl(bibleData, bibleVersion, bookNum, chapterStart, verseStart, verseEnd, chapterEnd);

    searchResultAstext.setLinkUrl(searchResultStart, searchResultStart + matchResult[0].length - 1, url);





    // Find the next match
    searchResult = searchElement.findText(searchRegex, searchResult);
    
  }; // END: Cycle through each Bible reference found

}; // END: function bibleParse(bibleData, bookNum, bookName, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd)


function getUrl(bibleData, bibleVersion, bookNum, chapterStart, verseStart, verseEnd, chapterEnd) {

  // Pull default Bible data if nothing was provided
  if ( ! bibleData ) bibleData = JSON.parse(UrlFetchApp.fetch(BIBLE_DATA_DEFAULT));

  // Set bibleVersion to default if not found in Bible data
  let bibleVersions = Object.keys(bibleData.bibleVersions);
  if ( ! bibleVersions.includes(bibleVersion) || bibleVersion == "default"  ) bibleVersion = bibleData.bibleVersions.default;

  // Get bookNames from bookNum
  var bookNameAbbr1 = bibleData.bookNames[bookNum][0];
  var bookNameAbbr2 = bibleData.bookNames[bookNum][1];
  var bookNameFull  = bibleData.bookNames[bookNum][bibleData.bookNames[bookNum].length - 1];

  // Convert any integers to string
  bookNum = bookNum.toString();
  chapterStart = chapterStart.toString();
  verseStart = verseStart.toString();
  if ( ! verseEnd ) verseEnd = verseStart;
  if ( verseEnd !== verseStart ) verseEnd = verseEnd.toString();
  if ( ! chapterEnd ) chapterEnd = chapterStart;
  if ( chapterEnd !== chapterStart ) chapterEnd = chapterEnd.toString();

  // Format book numbers, chapters, and verses
  var targetLength, padString;
  
  if ( bibleData.bibleVersions[bibleVersion].padStart.bookNum ) {
    targetLength = bibleData.bibleVersions[bibleVersion].padStart.bookNum.targetLength;
    padString = bibleData.bibleVersions[bibleVersion].padStart.bookNum.padString;

    bookNum = bookNum.padStart(targetLength, padString);
  };
  
  if ( bibleData.bibleVersions[bibleVersion].padStart.chapter ) {
    targetLength = bibleData.bibleVersions[bibleVersion].padStart.chapter.targetLength;
    padString = bibleData.bibleVersions[bibleVersion].padStart.chapter.padString;

    chapterStart = chapterStart.padStart(targetLength, padString);
    chapterEnd = chapterEnd.padStart(targetLength, padString);
  };
  
  if ( bibleData.bibleVersions[bibleVersion].padStart.verse ) {
    targetLength = bibleData.bibleVersions[bibleVersion].padStart.verse.targetLength;
    padString = bibleData.bibleVersions[bibleVersion].padStart.verse.padString;

    verseStart = verseStart.padStart(targetLength, padString);
    verseEnd = verseEnd.padStart(targetLength, padString);
  };

  let url = bibleData.bibleVersions[bibleVersion].urlFormat;
  url = url
    .replace(/<<bookNum>>/g, bookNum)
    .replace(/<<chapterStart>>/g, chapterStart)
    .replace(/<<verseStart>>/g, verseStart)
    .replace(/<<verseEnd>>/g, verseEnd)
    .replace(/<<chapterEnd>>/g, chapterEnd)
    .replace(/<bookNameAbbr1>>/g, bookNameAbbr1)
    .replace(/<<bookNameAbbr2>>/g, bookNameAbbr2)
    .replace(/<<bookNameFull>>/g, bookNameFull);

  // Remove range if single verse scripture only
  if ( chapterStart === chapterEnd && verseStart === verseEnd ) url = url.replace(/-[0-9]+$|-[0-9]+:[0-9]+:[0-9]+$/, '');

  return url;

}; // END: function getUrl(bibleData, bibleVersion, bookNum, chapterStart, verseStart, verseEnd, chapterEnd)


function testFunctions() {
  main();
  // Logger.log(getUrl(null, "default", 1, 1, 1, 3));
};
