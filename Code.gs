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

loadGlobals();

function loadGlobals() {

  var BIBLE_DATA_SOURCES = {
    "default": "en_jw_maj",
    "en_jw_maj": {
      "displayName": "English (JW.org)",
      "url": "https://github.com/majal/bible-linker-google-docs/raw/linker-v2-commenter/bible-data/en_jw.json"
    }
  };

  // Get user's last used bibleVersions
  const userProperties = PropertiesService.getUserProperties();
  var bibleDataUrl = userProperties.getProperty('bibleDataUrl');
  var bibleVersions = userProperties.getProperty('bibleVersions');

  // If bibleVersions is not a proper JSON, set to null
  try {
    bibleVersions = JSON.parse(bibleVersions);
  } catch {
    bibleVersions = null;
  };

  // If there is no bibleDataUrl, set to default
  if ( ! bibleDataUrl ) {
    bibleDataUrl = BIBLE_DATA_SOURCES[BIBLE_DATA_SOURCES.default].url;
  };

  // If there is no last used bibleVersions ... 
  if ( ! bibleVersions ) {

    // Pull bibleVersions from default data source
    let bibleData = JSON.parse(UrlFetchApp.fetch(bibleDataUrl));
    
    // Load bibleVersions into array, except "default"
    bibleVersions = [];
    for ( let bibleVersion of Object.keys(bibleData.bibleVersions) ) {
      if ( bibleVersion == "default" ) continue;
      bibleVersions.splice(bibleVersions.length, 0, bibleVersion);
    };

  };

  // Generate function names for the dynamic menu
  for ( let i = 0; i < bibleVersions.length; i++ ) {
    let dynamicMenuBibleVersion = 'dynamicFunctionCall_' + bibleVersions[i];
    this[dynamicMenuBibleVersion] = function() { bibleLinker(bibleDataUrl, bibleVersions[i]); };
  };

}; // END: loadGlobals()


//////////////////
// Dynamic menu //
//////////////////

function createMenu() {

  // Get user's last used bibleVersion and bibleDataUrl
  const userProperties = PropertiesService.getUserProperties();
  var bibleVersion = userProperties.getProperty('bibleVersion');
  var bibleDataUrl = userProperties.getProperty('bibleDataUrl');

  // If there is no bibleDataUrl, set to default
  if ( ! bibleDataUrl ) {
    bibleDataUrl = BIBLE_DATA_SOURCES[BIBLE_DATA_SOURCES.default].url;
  };

  // Fetch bibleData
  var bibleData = JSON.parse(UrlFetchApp.fetch(bibleDataUrl));

  // Set bibleVersions to default if not found in Bible data
  let bibleVersions = Object.keys(bibleData.bibleVersions);
  if ( ! bibleVersions.includes(bibleVersion) ) bibleVersion = bibleData.bibleVersions.default;

  // Set lastest used Bible version to the menu
  let displayName = bibleData.bibleVersions[bibleVersion].displayName;

  // Set main menu item
  var ui = DocumentApp.getUi();
  var menu = ui.createMenu(bibleData.strings.menu.title)
    .addItem( bibleData.strings.menu.doLink + ' ' + displayName, "dynamicFunctionCall_" + bibleVersion );

  // Set BIBLE_VERSIONS submenu
  var menuChooseBibleVersion = ui.createMenu(bibleData.strings.menu.chooseBibleVersion);

  // Load dynamic values to BIBLE_VERSIONS submenu
  for (let bibleVersionDynamic of bibleVersions) {
    if ( bibleVersionDynamic == 'default' ) continue;
    
    let bibleVersionDisplayName = bibleData.bibleVersions[bibleVersionDynamic].displayName;
    dynamicMenuBibleVersions = 'dynamicFunctionCall_' + bibleVersionDynamic;

    let pointer = ( bibleVersion == bibleVersionDynamic ) ? '▸\u00a0\u00a0' : '\u00a0\u00a0\u00a0\u00a0';
    menuChooseBibleVersion.addItem(pointer + bibleVersionDisplayName, dynamicMenuBibleVersions);
    
  };

  // Create menu 
  menu
    .addSubMenu(menuChooseBibleVersion)
    .addSeparator()
    .addItem('📝  Study tools', 'bibleLinker')
    .addToUi();

};

////////////////////
// Core functions //
////////////////////

function bibleLinker(bibleDataUrl, bibleVersion) {

  // Fetch bibleData
  var bibleData = JSON.parse(UrlFetchApp.fetch(bibleDataUrl));

  // Load bibleVersions into array, except "default"
  var bibleVersions = [];
  for ( let bibleVersion of Object.keys(bibleData.bibleVersions) ) {
    if ( bibleVersion == "default" ) continue;
    bibleVersions.splice(bibleVersions.length, 0, bibleVersion);
  };

  // Save last used values to user preferences
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('bibleDataUrl', bibleDataUrl);
  userProperties.setProperty('bibleVersion', bibleVersion);
  userProperties.setProperty('bibleVersions', JSON.stringify(bibleVersions));

  // Initialize Google Document
  var doc = DocumentApp.getActiveDocument();

  // Initialize access to Docs UI
  var ui = DocumentApp.getUi();

  // Note if document has active selection
  var docSelection = doc.getSelection();
  
  // Get single-chapter Bible books
  var bookSingleChapters = bibleData.bookSingleChapters;

  // Get search RegExes
  var searchRegexMultiChapters = bibleData.regEx.search.multiChapters;
  var searchRegexSingleChapters = bibleData.regEx.search.singleChapters;

  // Get whitespace RegEx, set to default if null
  var ws = bibleData.regEx.whitespace;

  // Get error messages
  var errorMsgParserTitle = bibleData.strings.errorMessages.parserError.title;
  var errorMsgParserBefore = bibleData.strings.errorMessages.parserError.messageBefore;
  var errorMsgParserAfter = bibleData.strings.errorMessages.parserError.messageAfter;

  //Initialize variables to use in loops
  var bookNum, bookName, searchRegex;

  // Expand bookNames with spaces to RegEx whitespaces
  for (bookName of Object.values(bibleData.bookNames)) {
    for ( let i=0; i < bookName.length; i++ ) {
      if ( bookName[i].includes(' ') ) {
        let nbspName = bookName[i].replace(/ /g, ws);
        bookName.splice(i, 0, nbspName);
        i++;      
      };
    };
  };  

  // Get book numbers, process each
  for (bookNum of Object.keys(bibleData.bookNames)) {

    // Get book names, process each
    for (bookName of Object.values(bibleData.bookNames[bookNum])) {
      
      // Modify RegEx format if it is a single-chapter book
      if ( bookSingleChapters.includes(bookNum) ) {
        searchRegex = bookName + ws + searchRegexSingleChapters;
      } else {
        searchRegex = bookName + ws + searchRegexMultiChapters;
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
              bibleParse(bibleData, bibleVersion, bookNum, searchResult, searchElement, searchRegex);
            } else {
              bibleParse(bibleData, bibleVersion, bookNum, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd);
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
          
          bibleParse(bibleData, bibleVersion, bookNum, searchResult, searchElement, searchRegex);
        
        } catch {

          // Show text where search error occurred
          searchResultTextSlice = searchResult.getElement().asText().getText().slice(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive() + 1);
          // ui.alert(errorMsgParserTitle, errorMsgParserBefore + searchResultTextSlice + errorMsgParserAfter, ui.ButtonSet.OK);

        };

      }; // END: For each bookName, each selection or whole document, perform a search

    }; // END: Get book names, process each

  }; // END: Get book numbers, process each

  // Receate menu
  createMenu();

}; // END: function bibleLinker(bibleDataUrl, bibleVersion)


function bibleParse(bibleData, bibleVersion, bookNum, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd) {

  // Get whitespace RegEx, set to default if null
  var ws = bibleData.regEx.whitespace;

  // Variable(s) and constant(s)
  var bookSingleChapters = bibleData.bookSingleChapters;

  // Cycle through each Bible reference found
  while (searchResult != null) {
    
    // Get reference start and end offsets
    let searchResultStart = searchResult.getStartOffset();
    let searchResultEnd = searchResult.getEndOffsetInclusive();
    
    // Isolate reference text
    let searchResultAstext = searchResult.getElement().asText();
    let searchResultText = searchResultAstext.getText();
    let searchResultTextSlice = searchResultText.slice(searchResultStart, searchResultEnd + 1);
    
    //////////////////////////////////////////////////////////////////
    // Split Bible references into linkable parts: referenceSplits, //
    // saved as arrays within an array for individual processing    //
    //////////////////////////////////////////////////////////////////

    // Split to arrays for every semicolon then save to splitsSemicolon
    let splitsSemicolon = searchResultTextSlice.split(';');
    // Add back the removed semicolon(s), all array values except the last one
    for ( let i=0; i < splitsSemicolon.length-1; i++ ) {
      splitsSemicolon[i] += ';';
    };

    let referenceSplits = [];
    // For every splitsSemicolon generated ...
    for ( let i=0; i < splitsSemicolon.length; i++ ) {
      // Split to arrays within array for every comma, then save to referenceSplits
      referenceSplits.push(splitsSemicolon[i].split(','));
      // If comma split(s) occured ...
      if ( referenceSplits[i].length > 1 ) {

        // Run for all split array values except the last one
        for ( let j=0; j < referenceSplits[i].length-1; j++ ) {
          
          // If verses separated by commas are consecutive, join together as one reference
          let verseNow = parseInt(referenceSplits[i][j].match(/\d+$/g)[0], 10);
          let verseNext = parseInt(referenceSplits[i][j+1].match(/\d+;*$/g)[0].replace(';', ''), 10);
          if ( verseNow + 1 == verseNext ) {
            referenceSplits[i][j] += ',' + referenceSplits[i][j+1];
            referenceSplits[i].splice(j+1, 1);
          };

          // Add back the removed comma(s) if next array exist, except the last array value
          // Conditional IF added due to the splice above, moving the last array backward
          if (referenceSplits[i][j+1]) referenceSplits[i][j] += ',';

        };

      };

    };

    /////////////////////////////////////////////////////////////////////
    // Generate variables for the linker:                              //
    // chapterStart, verseStart, verseEnd, chapterEnd, referenceStart, //
    // and referenceEnd for every referenceSplit to pass on to linker  //
    /////////////////////////////////////////////////////////////////////

    let referenceStart = searchResultStart, referenceEnd;
    let chapterStart, verseStart, verseEnd, chapterEnd;

    // Get chapters in each referenceSplits
    for ( let i=0; i < referenceSplits.length; i++ ) {
      
      let chapters = referenceSplits[i][0].match(/\d+:/g);
      if ( chapters.length == 1 ) {
        chapterStart = chapterEnd = parseInt(chapters[0].replace(':', ''), 10);
      } else {
        chapterStart = parseInt(chapters[0].replace(':', ''), 10);
        chapterEnd = parseInt(chapters[1].replace(':', ''), 10);
      };

      // Get verses in each referenceSplits
      for ( let j=0; j < referenceSplits[i].length; j++ ) {

        if ( referenceSplits[i][j].includes(':') ) {
          verseStart = parseInt(referenceSplits[i][j].match(/:\d+/g)[0].replace(':', ''), 10);
        } else {
          let re1 = new RegExp('^' + ws + '\\d+', 'g');
          let re2 = new RegExp('^' + ws, 'g');
          verseStart = parseInt(referenceSplits[i][j].match(re1)[0].replace(re2, ''), 10);
        };

        verseEnd = parseInt(referenceSplits[i][j].match(/\d+[,;]*$/g)[0].replace(/[,;]$/g, ''), 10);

        // Determine referenceEnd, and linkable start and end
        referenceEnd = referenceStart + referenceSplits[i][j].length;
        let linkableStart = referenceStart + referenceSplits[i][j].search(/\d/);
        if ( i == 0 && j == 0 ) linkableStart = referenceStart;
        let linkableEnd = referenceStart + referenceSplits[i][j].search(/\d\D*$/);
        
        /////////////////////////////////////////////
        // This is where the actual linking occurs //
        /////////////////////////////////////////////

        // Only set links if:
        // (1) there is no selection (null)
        // (2) or full line is selected
        // (2) or searchResult is within selection range // +1 is a quirk of .getEndOffsetInclusive()
        if ((!searchElementStart || !searchElementEnd)
        || (searchElementStart == -1 && searchElementEnd == -1)
        || (searchResultStart >= searchElementStart && searchResultStart + referenceSplits[i][j].length <= searchElementEnd + 1)) {

          let url = getUrl(bibleData, bibleVersion, bookNum, chapterStart, verseStart, verseEnd, chapterEnd);
          searchResultAstext.setLinkUrl(linkableStart, linkableEnd, url);
        
        };

        // Set referenceStart for the next iteration
        referenceStart = referenceEnd;

      };

    }; // END: Generate variables for the linker

    // Find the next match
    searchResult = searchElement.findText(searchRegex, searchResult);
    
  }; // END: Cycle through each Bible reference found

}; // END: function bibleParse(bibleData, bibleVersion, bookNum, searchResult, searchElement, searchRegex, searchElementStart, searchElementEnd)


function getUrl(bibleData, bibleVersion, bookNum, chapterStart, verseStart, verseEnd, chapterEnd) {

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


//////////////////////
// Helper functions //
//////////////////////

function onInstall(e) {
  onOpen(e);
};

function onOpen(e) {
  createMenu();
};
