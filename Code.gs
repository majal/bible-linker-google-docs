/* ***********************************************************************************
 * 
 *  Bible linker for Google Docs
 * 
 *  A Google Documents Apps Script that searches for Bible verses and creates links to
 *  a selection of online Bible sources.
 *
 *  For more information, visit: https://github.com/majal/bible-linker-google-docs
 *
 *  v2.0.0-beta-1.3.0
 * 
 *********************************************************************************** */

////////////////////////////////////////////////////////////
// Global variables and function, needed for dynamic menu //
////////////////////////////////////////////////////////////

const BIBLE_DATA_SOURCES = {
  "default": "en_jw",
  "en_jw": {
    "displayName": "English (JW.org)",
    "url": [
      "https://github.com/majal/bible-linker-google-docs/raw/main/bible-data/en_jw.json",
      "https://pastebin.com/raw/0W8738GK"
    ],
    "bibleVersions": [
      "en_jw_nwt",
      "en_jw_nwt_wol",
      "en_jw_nwtsty",
      "en_jw_nwtsty_wol",
      "en_jw_nwtrbi8",
      "en_jw_nwtrbi8_wol",
      "en_jw_kjv",
      "en_jw_kjv_wol",
      "en_jw_by",
      "en_jw_by_wol",
      "en_jw_asv",
      "en_jw_asv_wol",
      "en_jw_ebr",
      "en_jw_ebr_wol",
      "en_jw_int",
      "en_jw_int_wol"
    ],
    "strings": {
      "activate": {
        "appTitle": "Bible Linker",
        "activationItem": "Activate Bible Linker",
        "activationTitle": "Bible Linker Enabled",
        "activationMsg": "You may now use Bible Linker. Please navigate again to the menu to use it."
      },
      "errors": {
        "nullBibleData": "No Bible data available",
        "downloadJSON": {
          "title": "Failed to get data source",
          "messageBeforeSingular": "There was a problem fetching the following data source:\n\n",
          "messageBeforePlural": "There were problems fetching the following data sources:\n\n"
        }
      }
    }
  },
  "custom": {
    "displayName": "Custom data source",
    "strings": {
      "activate": {
        "appTitle": "Bible Linker",
        "activationItem": "Activate Bible Linker",
        "activationTitle": "Bible Linker Enabled",
        "activationMsg": "You may now use Bible Linker. Please navigate again to the menu to use it."
      },
      "errors": {
        "nullBibleData": "No custom data available",
        "downloadJSON": {
          "title": "Failed to get data source",
          "messageBeforeSingular": "There was a problem fetching the following data source:\n\n",
          "messageBeforePlural": "There were problems fetching the following data sources:\n\n"
        }
      }
    }
  }
};

//////////////////
// Dynamic menu //
//////////////////

dynamicMenuGenerate();

function dynamicMenuGenerate() {

  let bibleDataSource, bibleVersions;

  // Try-catch to check if PropertiesService is available, i.e. ScriptApp.AuthMode.NONE
  try {

    const userProperties = PropertiesService.getUserProperties();
    bibleDataSource = userProperties.getProperty('bibleDataSource');
    bibleVersions   = userProperties.getProperty('bibleVersions');

    // Error mitigation in case JSON from PropertiesService is malformed
    try {
      bibleVersions = JSON.parse(bibleVersions);
    } catch {
      bibleVersions = BIBLE_DATA_SOURCES[BIBLE_DATA_SOURCES.default].bibleVersions;
    };

  // If ScriptApp.AuthMode.NONE
  } catch {

    bibleDataSource = BIBLE_DATA_SOURCES.default;
    bibleVersions   = BIBLE_DATA_SOURCES[BIBLE_DATA_SOURCES.default].bibleVersions;

  };

  // Generate bibleVersion function names for the dynamic menu
  for ( let i = 0; i < bibleVersions.length; i++ ) {
    var dynamicMenuBibleVersion = 'dynamicFunctionCall_ver_' + bibleDataSource + bibleVersions[i];
    this[dynamicMenuBibleVersion] = function() { bibleLinker(bibleDataSource, bibleVersions[i]); };
  };

  // Generate bibleDataSource function names for the dynamic menu
  for ( let dsk of Object.keys(BIBLE_DATA_SOURCES) ) {
    if ( dsk == 'default' ) continue;
    var dynamicMenuBibleDataSource = 'dynamicFunctionCall_src_' + dsk;
    this[dynamicMenuBibleDataSource] = function() { chooseDataSource(dsk); };
  };

}; // END: dynamicMenuGenerate()


function createMenu() {

  // Get user's last used bibleDataSource and bibleVersion
  const userProperties = PropertiesService.getUserProperties();
  let bibleDataSource = userProperties.getProperty('bibleDataSource');
  let bibleVersion = userProperties.getProperty('bibleVersion');

  // If there is no bibleDataSource
  // or bibleDataSource not included in current list (keys)
  // then set to default value
  if ( ! bibleDataSource
  || ! Object.keys(BIBLE_DATA_SOURCES).includes(bibleDataSource) ) {
    bibleDataSource = BIBLE_DATA_SOURCES.default;
  };

  // Fetch bibleData from external source, throw error if bibleData is null
  let bibleData = bibleDataSource == 'custom' ? getBibleDataCustom() : getBibleData(BIBLE_DATA_SOURCES[bibleDataSource].url, bibleDataSource);
  if ( ! bibleData ) throw new Error(BIBLE_DATA_SOURCES[bibleDataSource].strings.errors.nullBibleData);

  // Set bibleVersion to default if not found in Bible data
  if ( ! Object.keys(bibleData.bibleVersions).includes(bibleVersion) ) bibleVersion = bibleData.bibleVersions.default;

  // Get needed strings and values
  let displayName        = bibleData.bibleVersions[bibleVersion].displayName;
  let selectorSelected   = bibleData.strings.menu.selector.selected;
  let selectorUnselected = bibleData.strings.menu.selector.unselected;
  let lengthLimit        = bibleData.strings.menu.lengthLimit;

  // Set main menu item
  var ui = DocumentApp.getUi();
  var menu = ui.createAddonMenu()
    .addItem( bibleData.strings.menu.doLink + ' ' + displayName, 'dynamicFunctionCall_ver_' + bibleDataSource + bibleVersion )
    .addSeparator();

  // Set bibleVersions submenu
  var menuChooseBibleVersion = ui.createMenu(bibleData.strings.menu.chooseBibleVersion);

  // Load dynamic values to bibleVersions submenu
  for ( let bibleVersionDynamic of Object.keys(bibleData.bibleVersions) ) {

    if ( bibleVersionDynamic == 'default' ) continue;
    
    let bibleVersionDisplayName = bibleData.bibleVersions[bibleVersionDynamic].displayName;

    // If bibleVersionDisplayName is over lengthLimit, truncate and add ellipsis …
    bibleVersionDisplayName = bibleVersionDisplayName.length > lengthLimit ? bibleVersionDisplayName.substring(0, lengthLimit) + '\u00a0…' : bibleVersionDisplayName;

    // Generate function names for dynamic submenu
    dynamicMenuBibleVersions = 'dynamicFunctionCall_ver_' + bibleDataSource + bibleVersionDynamic;

    // Add pointer at the beginning of the selected menu item
    let pointer = bibleVersion == bibleVersionDynamic ? selectorSelected : selectorUnselected;

    // Add menu item
    menuChooseBibleVersion.addItem(pointer + bibleVersionDisplayName, dynamicMenuBibleVersions);
    
  };

  // Set bibleDataSource submenu
  var menuChooseBibleDataSource = ui.createMenu(bibleData.strings.menu.chooseDataSource);

  // Load dynamic values to bibleDataSources submenu
  for ( let bibleDataSourceDynamic of Object.keys(BIBLE_DATA_SOURCES) ) {

    if ( bibleDataSourceDynamic == 'default' ) continue;

    let bibleDataSourceDisplayName = BIBLE_DATA_SOURCES[bibleDataSourceDynamic].displayName;

    // Only show custom data source if actually present
    if ( bibleDataSourceDynamic == 'custom' ) {

      let customBibleData = userProperties.getProperty('customBibleData');
      if ( ! customBibleData ) continue;

      bibleDataSourceDisplayName = 'Custom: ' + customBibleData;

    };

    // If bibleDataSourceDisplayName is over lengthLimit, truncate and add ellipsis …
    bibleDataSourceDisplayName = bibleDataSourceDisplayName.length > lengthLimit ? bibleDataSourceDisplayName.substring(0, lengthLimit) + '\u00a0…' : bibleDataSourceDisplayName;
    
    // Generate function names for dynamic submenu
    dynamicMenuBibleDataSource = 'dynamicFunctionCall_src_' + bibleDataSourceDynamic;

    // Add pointer at the beginning of the selected menu item
    let pointer = bibleDataSource == bibleDataSourceDynamic ? selectorSelected : selectorUnselected;

    // Add menu item
    menuChooseBibleDataSource.addItem(pointer + bibleDataSourceDisplayName, dynamicMenuBibleDataSource);
    
  };

  // Add custom entry at the end of submenu
  menuChooseBibleDataSource
    .addSeparator()
    .addItem(bibleData.strings.menu.customDataSource, 'chooseDataSourceCustom');

  // Get studyToolsDisplayName
  let studyToolsDisplayName = bibleData.html.studyTools.displayName;

  // Create menu 
  menu
    .addSubMenu(menuChooseBibleVersion)
    .addSubMenu(menuChooseBibleDataSource)
    .addSeparator()
    .addItem(studyToolsDisplayName, 'studyTools')
    .addToUi();

}; // END: function createMenu()


////////////////////
// Core functions //
////////////////////

function bibleLinker(bibleDataSource, bibleVersion) {

  // If there is no bibleDataSource
  // or bibleDataSource not included in current list (keys)
  // then set to default value
  if ( ! bibleDataSource
  || ! Object.keys(BIBLE_DATA_SOURCES).includes(bibleDataSource) ) {
    bibleDataSource = BIBLE_DATA_SOURCES.default;
  };

  // Fetch bibleData from external source, throw error if bibleData is null
  let bibleData = bibleDataSource == 'custom' ? getBibleDataCustom() : getBibleData(BIBLE_DATA_SOURCES[bibleDataSource].url, bibleDataSource);
  if ( ! bibleData ) throw new Error(BIBLE_DATA_SOURCES[bibleDataSource].strings.errors.nullBibleData);

  // Set bibleVersion to default if not found in Bible data
  if ( ! Object.keys(bibleData.bibleVersions).includes(bibleVersion) ) bibleVersion = bibleData.bibleVersions.default;

  // Load bibleVersions into array, except 'default'
  var bibleVersions = [];
  for ( let bvk of Object.keys(bibleData.bibleVersions) ) {
    if ( bvk == 'default' ) continue;
    bibleVersions.splice(bibleVersions.length, 0, bvk);
  };

  // Save last used values to user preferences
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('bibleDataSource', bibleDataSource);
  userProperties.setProperty('bibleVersion', bibleVersion);
  userProperties.setProperty('bibleVersions', JSON.stringify(bibleVersions));

  // Initialize Google Document
  var doc = DocumentApp.getActiveDocument();

  // Access Docs UI
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
    let bookNameLengthInit = bookName.length;
    for ( let i=0; i < bookNameLengthInit; i++ ) {
      if ( bookName[i].includes(' ') ) {
        let nbspName = bookName[i].replace(/ /g, ws);
        // Replace only if next name is different than previous name
        if ( nbspName != bookName[bookName.length-1] ) bookName.splice(bookName.length, 0, nbspName);
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
            ui.alert(errorMsgParserTitle, errorMsgParserBefore + searchResultTextSlice + errorMsgParserAfter, ui.ButtonSet.OK);

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
          ui.alert(errorMsgParserTitle, errorMsgParserBefore + searchResultTextSlice + errorMsgParserAfter, ui.ButtonSet.OK);

        };

      }; // END: For each bookName, each selection or whole document, perform a search

    }; // END: Get book names, process each

  }; // END: Get book numbers, process each

  // Receate menu
  createMenu();

}; // END: function bibleLinker(bibleDataSource, bibleVersion)


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
  var bookNameFull  = bibleData.bookNames[bookNum][2];

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


/////////////////////
// Other functions //
/////////////////////

function getBibleData(bibleDataSourceUrl, bibleDataSource) {

  // If there is no bibleDataSource
  // or bibleDataSource not included in current list (keys)
  // then set to default value
  if ( ! bibleDataSource
  || ! Object.keys(BIBLE_DATA_SOURCES).includes(bibleDataSource) ) {
    bibleDataSource = BIBLE_DATA_SOURCES.default;
  };

  // Get error messages
  let title                 = BIBLE_DATA_SOURCES[bibleDataSource].strings.errors.downloadJSON.title;
  let messageBeforeSingular = BIBLE_DATA_SOURCES[bibleDataSource].strings.errors.downloadJSON.messageBeforeSingular;
  let messageBeforePlural   = BIBLE_DATA_SOURCES[bibleDataSource].strings.errors.downloadJSON.messageBeforePlural;

  // Access Docs UI
  var ui = DocumentApp.getUi();

  // Define variables
  let bibleData, bibleDataJSON;

  // If data source contains multiple URLs, isArray
  if ( Array.isArray(bibleDataSourceUrl) ) {

    // For each URL in array, exit once a proper JSON is found
    for ( let i = 0; i < bibleDataSourceUrl.length; i++ ) {

      // Try to download JSON
      try {
      
        bibleData = UrlFetchApp.fetch(bibleDataSourceUrl[i]);
      
      // Continue with next URL in array if URL is invalid
      } catch {

        continue;
      
      };

      // If JSON exists and downloadable
      if ( bibleData.getResponseCode() == 200 ) {

        // Return and exit if valid JSON
        try {

          bibleDataJSON = JSON.parse(bibleData.getContentText());
          return bibleDataJSON;

        // Continue with next URL in array if JSON is invalid
        } catch {

          continue;
        
        };

      };

    };

    // If no valid JSON is found among all the URLs
    if ( ! bibleDataJSON ) {

      // Send alert to UI and return null
      ui.alert(title, messageBeforePlural + bibleDataSourceUrl.join('\n'), ui.ButtonSet.OK);
      return null;

    };

  // If data source contain only a single URL string
  } else {

    // Try to download JSON
    try {
    
      bibleData = UrlFetchApp.fetch(bibleDataSourceUrl);
    
    // If URL invalid, send alert to UI and return null
    } catch {

      ui.alert(title, messageBeforeSingular + bibleDataSourceUrl, ui.ButtonSet.OK);
      return null;
    
    };
    
    // Return and exit if valid JSON
    try {

      bibleDataJSON = JSON.parse(bibleData.getContentText());
      return bibleDataJSON;
    
    // If not a valid JSON, send alert to UI and return null
    } catch {

      ui.alert(title, messageBeforeSingular + bibleDataSourceUrl, ui.ButtonSet.OK);
      return null;

    };
  
  };

}; // END: getBibleData(bibleDataSourceUrl, bibleDataSource)


function getBibleDataCustom() {

  const userProperties = PropertiesService.getUserProperties();
  let customBibleData = userProperties.getProperty('customBibleData');

  // Try if customBibleData is a JSON object
  try {

    let url = JSON.parse(customBibleData).url;

    // If JSON contains URL
    if ( url ) {

      return getBibleData(url, 'custom');
    
    // If JSON does not contains URL, reset to default
    } else {

      userProperties.setProperty('bibleDataSource', BIBLE_DATA_SOURCES.default);
      userProperties.deleteProperty('customBibleData');

      createMenu();
    
    };
  
  // If customBibleData is not JSON
  } catch {
  
    return getBibleData(customBibleData, 'custom');
  
  };

};


function chooseDataSource(bibleDataSource) {
  
  // If there is no bibleDataSource
  // or bibleDataSource not included in current list (keys)
  // then set to default value
  if ( ! bibleDataSource
  || ! Object.keys(BIBLE_DATA_SOURCES).includes(bibleDataSource) ) {
    bibleDataSource = BIBLE_DATA_SOURCES.default;
  };

  // Fetch bibleData from external source, throw error if bibleData is null
  let bibleData = bibleDataSource == 'custom' ? getBibleDataCustom() : getBibleData(BIBLE_DATA_SOURCES[bibleDataSource].url, bibleDataSource);
  if ( ! bibleData ) throw new Error(BIBLE_DATA_SOURCES[bibleDataSource].strings.errors.nullBibleData);

  let updateTitle         = bibleData.strings.bibleDataSource.update.title;
  let updateMessageBefore = bibleData.strings.bibleDataSource.update.messageBefore;
  let updateMessageAfter  = bibleData.strings.bibleDataSource.update.messageAfter;

  // Set bibleDataSource to user preferences
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('bibleDataSource', bibleDataSource);

  // Recreate menu
  createMenu();

  // Inform user of change of data source
  var ui = DocumentApp.getUi();
  ui.alert(updateTitle, updateMessageBefore + BIBLE_DATA_SOURCES[bibleDataSource].displayName + updateMessageAfter, ui.ButtonSet.OK);

};


function chooseDataSourceCustom() {

  // Get user's last used bibleDataSource
  const userProperties = PropertiesService.getUserProperties();
  let bibleDataSource = userProperties.getProperty('bibleDataSource');

  // If there is no bibleDataSource
  // or bibleDataSource not included in current list (keys)
  // then set to default value
  if ( ! bibleDataSource
  || ! Object.keys(BIBLE_DATA_SOURCES).includes(bibleDataSource) ) {
    bibleDataSource = BIBLE_DATA_SOURCES.default;
  };

  // Fetch bibleData from external source, throw error if bibleData is null
  let bibleData = bibleDataSource == 'custom' ? getBibleDataCustom() : getBibleData(BIBLE_DATA_SOURCES[bibleDataSource].url, bibleDataSource);
  if ( ! bibleData ) throw new Error(BIBLE_DATA_SOURCES[bibleDataSource].strings.errors.nullBibleData);

  // Get strings
  let inputTitle     = bibleData.strings.customDataSource.input.title;
  // let inputMsgBefore = bibleData.strings.customDataSource.input.messageBefore;
  let inputMsgAfter  = bibleData.strings.customDataSource.input.messageAfter;
  let errorTitle     = bibleData.strings.customDataSource.error.title;
  let errorMessage   = bibleData.strings.customDataSource.error.message;
  let successTitle     = bibleData.strings.customDataSource.success.title;
  let successMsgBefore = bibleData.strings.customDataSource.success.messageBefore;

  // Access Docs UI
  var ui = DocumentApp.getUi();

  // Initialize variables
  let customBibleDataJSON;

  // Get custom JSON or URL
  let response = ui.prompt(inputTitle, inputMsgAfter, ui.ButtonSet.OK_CANCEL);

  // Exit if there is no input
  if ( response.getResponseText().length == 0 ) return;

  // If the user clicked the OK button
  if ( response.getSelectedButton() == ui.Button.OK ) {

    // Check if input is valid JSON
    try {

      customBibleDataJSON = JSON.parse(response.getResponseText());

      // If JSON is valid ...
      if ( customBibleDataJSON ) {

        // Check if URL(s) point to valid JSON
        if ( getBibleData(customBibleDataJSON.url, bibleDataSource) ) {

          // Upload to userProperties
          try {

            userProperties.setProperty('bibleDataSource', 'custom');
            userProperties.setProperty('customBibleData', JSON.stringify(customBibleDataJSON));

            ui.alert(successTitle, successMsgBefore + JSON.stringify(customBibleDataJSON, null, '\u00a0\u00a0'), ui.ButtonSet.OK);
            
            createMenu();            
            return;
          
          // Catch if userProperties.setProperty() returns an error
          } catch(e) {

            // Notify about the error
            ui.alert(errorTitle, errorMessage + e + '\n\n' + JSON.stringify(customBibleDataJSON, null, '\u00a0\u00a0'), ui.ButtonSet.OK);

            // Restart function
            chooseDataSourceCustom();
            return;
          
          };

        // If URL(s) in JSON is invalid
        } else {

          // Notify about the error
          // ui.alert(errorTitle, errorMessage + e + '\n\n' + JSON.stringify(customBibleDataJSON, null, '\u00a0\u00a0'), ui.ButtonSet.OK);

          // Restart function
          chooseDataSourceCustom();
          return;

        }

      };

    // If not valid JSON, try if it is a URL pointing to a valid JSON
    } catch {

      // If URL is valid, upload URL to userProperties
      if ( getBibleData(response.getResponseText(), bibleDataSource) ) {

        userProperties.setProperty('bibleDataSource', 'custom');
        userProperties.setProperty('customBibleData', response.getResponseText());

        ui.alert(successTitle, successMsgBefore + response.getResponseText(), ui.ButtonSet.OK);

        createMenu();
        return;
      
      // If URL is invalid
      } else {
    
        // Notify about the error
        // ui.alert(errorTitle, errorMessage + response.getResponseText(), ui.ButtonSet.OK);

        // Restart function
        chooseDataSourceCustom();
        return;

      };

    }; // END: Check if input is valid JSON

  } else {
    
    return;

  }; // END: If the user clicked the OK button

}; // END: chooseDataSourceCustom()


function studyTools() {

  // Get user's last used bibleDataSource
  const userProperties = PropertiesService.getUserProperties();
  let bibleDataSource = userProperties.getProperty('bibleDataSource');

  // If there is no bibleDataSource
  // or bibleDataSource not included in current list (keys)
  // then set to default value
  if ( ! bibleDataSource
  || ! Object.keys(BIBLE_DATA_SOURCES).includes(bibleDataSource) ) {
    bibleDataSource = BIBLE_DATA_SOURCES.default;
  };

  // Fetch bibleData from external source, throw error if bibleData is null
  let bibleData = bibleDataSource == 'custom' ? getBibleDataCustom() : getBibleData(BIBLE_DATA_SOURCES[bibleDataSource].url, bibleDataSource);
  if ( ! bibleData ) throw new Error(BIBLE_DATA_SOURCES[bibleDataSource].strings.errors.nullBibleData);
  
  // Fetch studyTools HTML content
  let htmlContent = UrlFetchApp.fetch(bibleData.html.studyTools.url);
  let htmlOutput = HtmlService.createHtmlOutput(htmlContent);

  // Show studyTools
  DocumentApp.getUi().showModalDialog(htmlOutput, bibleData.html.studyTools.windowLabel);

};


function activateAddon() {

  createMenu();

  // Access Docs UI
  var ui = DocumentApp.getUi();

  ui.alert(BIBLE_DATA_SOURCES[BIBLE_DATA_SOURCES.default].strings.activate.activationTitle,
    BIBLE_DATA_SOURCES[BIBLE_DATA_SOURCES.default].strings.activate.activationMsg,
    ui.ButtonSet.OK);  

};

//////////////////////
// Helper functions //
//////////////////////

function onInstall(e) {

  onOpen(e);

};


function onOpen(e) {

  // Access Docs UI
  var ui = DocumentApp.getUi();
  
  // If AuthMode not FULL, create temporary menu
  if (e && e.authMode != ScriptApp.AuthMode.FULL) {

    var menu = ui.createMenu(BIBLE_DATA_SOURCES[BIBLE_DATA_SOURCES.default].strings.activate.appTitle);
    menu
      .addItem(BIBLE_DATA_SOURCES[BIBLE_DATA_SOURCES.default].strings.activate.activationItem, 'activateAddon')
      .addToUi();

  // If ScriptApp.AuthMode is FULL, activated when passed from onInstall(e)
  // https://developers.google.com/apps-script/add-ons/concepts/editor-auth-lifecycle
  } else {

    createMenu();

  };

};
