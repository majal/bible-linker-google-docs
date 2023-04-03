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

function main(bibleDataSelection, linkFormat) {

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
              bibleParse(bibleData, bookNum, searchResult, searchElement, searchRegex);
            } else {
              bibleParse(bibleData, bookNum, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd);
            };
            
          } catch {

            // Show text where search error occurred
            searchResultTextSlice = searchResult.getElement().asText().getText().slice(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive() + 1);
            ui.alert(errorMsgParserTitle, errorMsgParserBefore + searchResultTextSlice + errorMsgParserAfter, ui.ButtonSet.OK);

          };

        };

      // Perform search on whole document
      } else {

        searchElement = doc.getBody();
        searchResult = searchElement.findText(searchRegex);
        
        // Send found matches to parser
        try {

          bibleParse(bibleData, bookNum, searchResult, searchElement, searchRegex);
        
        } catch {

          // Show text where search error occurred
          searchResultTextSlice = searchResult.getElement().asText().getText().slice(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive() + 1);
          ui.alert(errorMsgParserTitle, errorMsgParserBefore + searchResultTextSlice + errorMsgParserAfter, ui.ButtonSet.OK);

        };

      }; // END: For each bookName, each selection or whole document, perform a search

    }; // END: Get book names, process each

  }; // END: Get book numbers, process each

}; // END: function main(bibleDataSelection)


function bibleParse(bibleData, bookNum, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd) {

  // Pull default Bible data if nothing was provided
  if ( ! bibleData ) bibleData = JSON.parse(UrlFetchApp.fetch(BIBLE_DATA_DEFAULT));

  // Variable(s) and constant(s)
  var bookSingleChapters = bibleData.bookSingleChapters;

  // Cycle through each Bible reference found
  while (searchResult != null) {

    // Set reference start and end
    var searchResultStart = searchResult.getStartOffset();
    var searchResultEnd = searchResult.getEndOffsetInclusive();

    // Isolate reference only
    var searchResultAstext = searchResult.getElement().asText();
    var searchResultText = searchResultAstext.getText();
    var searchResultTextSlice = searchResultText.slice(searchResultStart, searchResultEnd + 1);

    Logger.log('|' + searchResultTextSlice + '|');

    /*
    // Split at semicolon (;)
    var bibleref_split = searchResultTextSlice.split(';');
  
    // Retain verses only, remove if it does not contain colon (:), exception on single chapter books
    for (let n=0; n < bibleref_split.length; n++) {
      if (!single_chapters.includes(bookNum) && !bibleref_split[n].includes(':')) {
        bibleref_split.splice(n, 1);
        n--;
      }
    }

    // Split by comma (,)
    for (let n=0; n < bibleref_split.length; n++) {
      if (bibleref_split[n].includes(',')) {
        bibleref_split[n] = bibleref_split[n].split(',');

        // Rejoin if verses are consecutive
        if (Array.isArray(bibleref_split[n])) {
          for (let m=1; m < bibleref_split[n].length; m++) {
            if (parseInt(bibleref_split[n][m-1].match(/[0-9]+$/), 10) + 1 == parseInt(bibleref_split[n][m], 10)) {
              bibleref_split[n][m-1] += ',' + bibleref_split[n][m];
              bibleref_split[n].splice(m, 1);
              m--;
            }
          }

          // Convert array to string if consecutive verses
          if (bibleref_split[n].length == 1) {
            bibleref_split[n] = bibleref_split[n][0].toString();
          }
        }
      }
    }
    */

    /*
    // Initialize vars and offsets
    let select_start = 0;
    let select_end = 0;
    let url = '';
    let offset_reference = 0;
    
    // Parse and process
    for (let n=0; n < bibleref_split.length; n++) {

      // Declare variable(s)
      let chapters = [], verseStart, verseEnd;

      if (Array.isArray(bibleref_split[n])) {
        for (let m=0; m < bibleref_split[n].length; m++) {

          // Get chapter(s)
          if (single_chapters.includes(bookNum)) {
            chapters[0] = 1;
            chapters[1] = 1;
          } else {
            chapters = bibleref_split[n][0].match(/[0-9]+:/g);
            if (chapters.length == 1) {
              chapters[0] = chapters[0].replace(':', '');
              chapters[1] = chapters[0];
            } else {
              chapters[0] = chapters[0].replace(':', '');
              chapters[1] = chapters[1].replace(':', '');
            }
          }

          // Get verse(s)
          if (bibleref_split[n][m].includes(':')) {
            verseStart = bibleref_split[n][m].match(/:[0-9]+/).toString().replace(':', '');
            verseEnd = bibleref_split[n][m].match(/[0-9]+\s*$/).toString().replace(':', '');
          } else {
            verseStart = bibleref_split[n][m].match(/\s[0-9]+/).toString();
            verseEnd = bibleref_split[n][m].match(/[0-9]+\s*$/).toString();
          }

          // Get url link
          url = get_url(bible_version, bookNum, chapters[0], chapters[1], verseStart, verseEnd);

          // Get url text ranges
          let url_text_len = bibleref_split[n][m].trim().length;
          select_start = searchResultStart + offset_reference;
          select_end = select_start + url_text_len - 1;
          
          // Set links if there is no selection or if selection exists and it is within the selection
          if ((!searchElementStart && !searchElementEnd) || (searchElementStart == -1 && searchElementEnd == -1) || (select_start >= searchElementStart && select_end <= searchElementEnd)) {
            searchResultAstext.setLinkUrl(select_start, select_end, url);
          }
          
          // Add to reference offset, plus two for comma/colon and space
          offset_reference += url_text_len + 2;
        }

      } else {

        // Get chapter(s)
        if (single_chapters.includes(bookNum)) {
          chapters[0] = 1;
          chapters[1] = 1;
        } else {
          chapters = bibleref_split[n].match(/[0-9]+:/g);
          if (chapters.length == 1) {
            chapters[0] = chapters[0].replace(':', '');
            chapters[1] = chapters[0];
          } else {
            chapters[0] = chapters[0].replace(':', '');
            chapters[1] = chapters[1].replace(':', '');
          }
        }

        // Get verse(s)
        if (single_chapters.includes(bookNum)) {
          verseStart = bibleref_split[n].match(/\s[0-9]+/).toString();
          verseEnd = bibleref_split[n].match(/[0-9]+\s*$/).toString();
        } else {
          verseStart = bibleref_split[n].match(/:[0-9]+/).toString().replace(':', '');
          verseEnd = bibleref_split[n].match(/[0-9]+\s*$/).toString().replace(':', '');
        }

        // Get url link
        url = get_url(bible_version, bookNum, chapters[0], chapters[1], verseStart, verseEnd);

        // Get url text ranges
        let url_text_len = bibleref_split[n].trim().length;
        select_start = searchResultStart + offset_reference;
        select_end = select_start + url_text_len - 1;
        
        // Set links if there is no selection or if selection exists and it is within the selection
        if ((!searchElementStart && !searchElementEnd) || (searchElementStart == -1 && searchElementEnd == -1) || (select_start >= searchElementStart && select_end <= searchElementEnd)) { 
          searchResultAstext.setLinkUrl(select_start, select_end, url);
        }

        // Add to reference offset, plus two for comma/colon and space
        offset_reference += url_text_len + 2

      };

    };
    */
    // Find the next match
    searchResult = searchElement.findText(searchRegex, searchResult);
    
  };

}; // END: function bibleParse(bibleData, bookNum, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd)


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
