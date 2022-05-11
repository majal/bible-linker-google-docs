/* ***********************************************************************************
 * 
 *  Bible linker for Google Docs
 * 
 *  A Google Documents Apps Script that searches for Bible verses and creates links to a selection of online Bible sources.
 * 
 *  For more information, visit: https://github.com/majal/bible-linker-google-docs
 *
 *********************************************************************************** */


function onInstall(e) {
  onOpen(e);
}


function onOpen(e) {
  create_menu();
}


function create_menu() {

  // Get lastest used Bible version by user
  const userProperties = PropertiesService.getUserProperties();
  var last_used_bible_version = userProperties.getProperty('last_used_bible_version');
  if (last_used_bible_version == null) last_used_bible_version = 'nwtsty_wol';

  // Pull supported Bible versions
  const bible_versions = consts('bible_versions');
  const bible_versions_keys = Object.keys(bible_versions);

  // Set lastest used Bible version to the menu
  let last_used_bible_label = bible_versions[last_used_bible_version];
  let dynamic_name_bible_linker = 'bible_linker_' + last_used_bible_version;
  let search_label = '🔗⠀Link verses using ' + last_used_bible_label;

  // Set menu
  var ui = DocumentApp.getUi();
  var menu = ui.createMenu('Bible Linker')
    .addItem(search_label, dynamic_name_bible_linker)

  // Set Bible version dynamic submenus
  var submenu_bible_ver = ui.createMenu('📖⠀Choose Bible version');
  for (let n=0; n < bible_versions_keys.length; n++) {
    let key = bible_versions_keys[n];
    let dynamic_name_bible_linker = 'bible_linker_' + key;
    let last_used_pointer = (last_used_bible_version == key) ? '▸ ⠀' : '⠀⠀';
    submenu_bible_ver.addItem(last_used_pointer + bible_versions[key], dynamic_name_bible_linker);
  }

  // Create menu 
  menu
    .addSubMenu(submenu_bible_ver)
    .addSeparator()
    .addItem('📝⠀Study tools', 'study_tools')
    .addToUi();

}


// Dynamic menu hack
const bible_versions = consts('bible_versions');
const bible_versions_keys = Object.keys(bible_versions);
for (let n=0; n < bible_versions_keys.length; n++) {
  let key = bible_versions_keys[n];
  let dynamic_name_bible_linker = 'bible_linker_' + key;
  this[dynamic_name_bible_linker] = function() { bible_linker(key); };
}


function study_tools() {
  var html_content = `
  <style>
    html {font-family: "Open Sans", Arial, sans-serif;}

    li {padding: 0 0 20px 0;}

    .button {
      background-color: #008CBA;
      border: 2px solid #008CBA;
      border-radius: 8px;
      font-weight: bold;
      color: #FFF;
      text-align: center;
      text-decoration: none;
      font-size: 16px;
      margin: 30px auto 10px auto;
      padding: 12px 24px;
      display:block;
      transition-duration: 0.4s;
      cursor: pointer;
    }

    .button:hover {
      box-shadow: 0 6px 16px 0 rgba(0,0,0,0.24), 0 9px 50px 0 rgba(0,0,0,0.19);
    }

    .button:active {
      box-shadow: 0 2px 50px 0 rgba(0,0,0,0.24), 0 5px 10px 0 rgba(0,0,0,0.19);
      transform: translateY(4px);
    }
  </style>
  
  <base target="_blank">

  <p>Tools to help you get a deeper understanding of the Bible:</p>

  <ul>
    <li><strong><a href="https://wol.jw.org/">Watchtower Online Library</a> (WOL)</strong> - A research tool to find explanatory articles about Bible verses and topics.</li>
    <li><strong><a href="https://www.jw.org/finder?docid=802013025">JW Library</a></strong> - Bible library in your pocket.</li>
    <li><strong><a href="https://www.jw.org/finder?docid=1011539">Study tools</a></strong> on <a href="https://www.jw.org/">jw.org</a>.</li>
  </ul>

  <input class="button" type="button" value="Got it!" onClick="google.script.host.close()" />
  `

  var htmlOutput = HtmlService
    .createHtmlOutput(html_content);
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Bible study tools');

}


function bible_linker(bible_version) {

  // Initialize Google Docs
  var doc = DocumentApp.getActiveDocument();

  // Set the latest used Bible version
  if (bible_version == undefined || bible_version == null) bible_version = 'nwtsty_wol';
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('last_used_bible_version', bible_version);
  create_menu();

  // Get names of Bible books
  var nwt_bookName = consts('nwt_bookName');
  var nwt_bookAbbrev1 = consts('nwt_bookAbbrev1');
  var nwt_bookAbbrev2 = consts('nwt_bookAbbrev2');

  // Run parser for each Bible name
  for (let n=0; n < nwt_bookName.length; n++) {
    bible_search(doc, bible_version, nwt_bookName[n], n+1);
  }
  for (let n=0; n < nwt_bookAbbrev1.length; n++) {
    bible_search(doc, bible_version, nwt_bookAbbrev1[n], n+1);
  }
  for (let n=0; n < nwt_bookAbbrev2.length; n++) {
    bible_search(doc, bible_version, nwt_bookAbbrev2[n], n+1);
  }

}


function bible_search(doc, bible_version, bible_name, bible_num) {

  // Set variables and constants
  var single_chapters = consts('single_chapter_bible_nums');
  var search_field, search_result;
  var selection = doc.getSelection();

  var err_msg_title = 'Oops!';
  var err_msg1 = 'There was an error processing this line:\n\n';
  var err_msg2 = "\n\nIs there a typo? (Tip: It's usually the spaces.)";

  // Search for Bible references
  if (single_chapters.includes(bible_num)) {
    var search_string = bible_name + ' [0-9 ,-]+';
  } else {
    var search_string = bible_name + ' [0-9]+:[0-9 ,;:-]+';
  }

  // Check if selection is present and only process that
  if (selection) {
    var range_elements = selection.getRangeElements();

    // Search within each selection element and pass results to parser 
    for (let n=0; n < range_elements.length; n++) {
      search_field = range_elements[n].getElement();
      search_result = search_field.findText(search_string);
      
      // Because the parser can hit unexpected errors with typos ;-)
      try {
        bible_parse(bible_version, bible_name, bible_num, search_result, search_field, search_string);
      } catch {
        var ui = DocumentApp.getUi();
        ui.alert(err_msg_title, err_msg1 + search_result.getElement().asText().getText() + err_msg2, ui.ButtonSet.OK);
      }

    }

  } else {

    // Perform search on whole document
    search_field = doc.getBody();
    search_result = search_field.findText(search_string);

    // Pass results to parser, and because the parser can hit unexpected errors with typos ;-)
    try {
      bible_parse(bible_version, bible_name, bible_num, search_result, search_field, search_string);
    } catch {
      var ui = DocumentApp.getUi();
      ui.alert(err_msg_title, err_msg1 + search_result.getElement().asText().getText() + err_msg2, ui.ButtonSet.OK);
    }

  } 

}


function bible_parse(bible_version, bible_name, bible_num, search_result, search_field, search_string) {

  // Variable(s) and constant(s)
  var single_chapters = consts('single_chapter_bible_nums');

  // Cycle through each Bible reference found
  while (search_result != null) {

    // Set reference start and end
    var search_result_start = search_result.getStartOffset();
    var search_result_end = search_result.getEndOffsetInclusive();

    // Isolate reference only
    var search_result_astext = search_result.getElement().asText();
    var search_result_text = search_result_astext.getText();
    var search_result_text_slice = search_result_text.slice(search_result_start, search_result_end + 1);

    // Split at semicolon (;)
    var bibleref_split = search_result_text_slice.split(';');
  
    // Retain verses only, remove if it does not contain colon (:), exception on single chapter books
    for (let n=0; n < bibleref_split.length; n++) {
      if (!single_chapters.includes(bible_num) && !bibleref_split[n].includes(':')) {
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

    // Initialize vars and offsets
    let select_start = 0;
    let select_end = 0;
    let url = '';
    let offset_reference = 0;

    // Parse and process
    for (let n=0; n < bibleref_split.length; n++) {

      // Declare variable(s)
      let chapters = [], verse_start, verse_end;

      if (Array.isArray(bibleref_split[n])) {
        for (let m=0; m < bibleref_split[n].length; m++) {

          // Get chapter(s)
          if (single_chapters.includes(bible_num)) {
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
            verse_start = bibleref_split[n][m].match(/:[0-9]+/).toString().replace(':', '');
            verse_end = bibleref_split[n][m].match(/[0-9]+$/).toString().replace(':', '');
          } else {
            verse_start = bibleref_split[n][m].match(/ [0-9]+/).toString();
            verse_end = bibleref_split[n][m].match(/[0-9]+$/).toString();
          }

          // Get url link
          url = get_url(bible_version, bible_num, chapters[0], chapters[1], verse_start, verse_end);

          // Get url text ranges
          let url_text_len = bibleref_split[n][m].trim().length;
          select_start = search_result_start + offset_reference;
          select_end = select_start + url_text_len - 1;
          
          // Set links
          search_result_astext.setLinkUrl(select_start, select_end, url);
          
          // Add to reference offset, plus two for comma/colon and space
          offset_reference += url_text_len + 2;
        }

      } else {

        // Get chapter(s)
        if (single_chapters.includes(bible_num)) {
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
        if (single_chapters.includes(bible_num)) {
          verse_start = bibleref_split[n].match(/ [0-9]+/).toString();
          verse_end = bibleref_split[n].match(/[0-9]+$/).toString();
        } else {
          verse_start = bibleref_split[n].match(/:[0-9]+/).toString().replace(':', '');
          verse_end = bibleref_split[n].match(/[0-9]+$/).toString().replace(':', '');
        }

        // Get url link
        url = get_url(bible_version, bible_num, chapters[0], chapters[1], verse_start, verse_end);

        // Get url text ranges
        let url_text_len = bibleref_split[n].trim().length;
        select_start = search_result_start + offset_reference;
        select_end = select_start + url_text_len - 1;
        
        // Set links
        search_result_astext.setLinkUrl(select_start, select_end, url);

        // Add to reference offset, plus two for comma/colon and space
        offset_reference += url_text_len + 2
      }
    }

    // Find the next match
    search_result = search_field.findText(search_string, search_result);
  }

}


function get_url(bible_version, bible_name_num, chapter_start, chapter_end, verse_start, verse_end) {

  // Declare variables
  var url_head, url_hash;

  // Initialize integer values
  chapter_start = parseInt(chapter_start, 10);
  chapter_end = parseInt(chapter_end, 10);
  verse_start = parseInt(verse_start, 10);
  verse_end = parseInt(verse_end, 10);

  // Choices from Bible versions
  // Make sure to add entries in function consts(bible_versions)
  switch (bible_version) {
  
    case 'nwt':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=nwt&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
      }
    break;

    case 'nwt_wol':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/nwt/';
      url_hash1 = '#s=';
      url_hash2 = '&study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;

    case 'nwtsty':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=nwtsty&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
      }
    break;

    case 'nwtsty_wol':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/nwtsty/';
      url_hash1 = '#s=';
      url_hash2 = '&study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;
      
    case 'nwtrbi8':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=Rbi8&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
      }
    break;

    case 'nwtrbi8_wol':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/Rbi8/';
      url_hash1 = '#s=';
      url_hash2 = '&study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;

    case 'kjv_jw':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=bi10&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
      }
    break;

    case 'kjw_wol':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/bi10/';
      url_hash1 = '#s=';
      url_hash2 = '&study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;

  case 'by_jw':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=by&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
      }
    break;

    case 'by_wol':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/by/';
      url_hash1 = '#s=';
      url_hash2 = '&study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;

case 'asv_jw':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=bi22&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
      }
    break;

    case 'asv_wol':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/bi22/';
      url_hash1 = '#s=';
      url_hash2 = '&study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;

    case 'ebr_jw':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=rh&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
      }
    break;

    case 'ebr_wol':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/rh/';
      url_hash1 = '#s=';
      url_hash2 = '&study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash1 + verse_start + url_hash2 + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;

    default:
      undefined;
  }

}


function consts(const_name) {
  switch (const_name) {

    case 'nwt_bookName':
      return ["Genesis", "Exodus", "Leviticus", "Numbers", "Deuteronomy", "Joshua", "Judges", "Ruth", "1 Samuel", "2 Samuel", "1 Kings", "2 Kings", "1 Chronicles", "2 Chronicles", "Ezra", "Nehemiah", "Esther", "Job", "Psalm", "Proverbs", "Ecclesiastes", "Song of Solomon", "Isaiah", "Jeremiah", "Lamentations", "Ezekiel", "Daniel", "Hosea", "Joel", "Amos", "Obadiah", "Jonah", "Micah", "Nahum", "Habakkuk", "Zephaniah", "Haggai", "Zechariah", "Malachi", "Matthew", "Mark", "Luke", "John", "Acts", "Romans", "1 Corinthians", "2 Corinthians", "Galatians", "Ephesians", "Philippians", "Colossians", "1 Thessalonians", "2 Thessalonians", "1 Timothy", "2 Timothy", "Titus", "Philemon", "Hebrews", "James", "1 Peter", "2 Peter", "1 John", "2 John", "3 John", "Jude", "Revelation"];
      break;

    case 'nwt_bookAbbrev1':
      return ["Ge", "Ex", "Le", "Nu", "De", "Jos", "Jg", "Ru", "1Sa", "2Sa", "1Ki", "2Ki", "1Ch", "2Ch", "Ezr", "Ne", "Es", "Job", "Ps", "Pr", "Ec", "Ca", "Isa", "Jer", "La", "Eze", "Da", "Ho", "Joe", "Am", "Ob", "Jon", "Mic", "Na", "Hab", "Zep", "Hag", "Zec", "Mal", "Mt", "Mr", "Lu", "Joh", "Ac", "Ro", "1Co", "2Co", "Ga", "Eph", "Php", "Col", "1Th", "2Th", "1Ti", "2Ti", "Tit", "Phm", "Heb", "Jas", "1Pe", "2Pe", "1Jo", "2Jo", "3Jo", "Jude", "Re"];
      break;

    case 'nwt_bookAbbrev2':
      return ["Gen.", "Ex.", "Lev.", "Num.", "Deut.", "Josh.", "Judg.", "Ruth", "1 Sam.", "2 Sam.", "1 Ki.", "2 Ki.", "1 Chron.", "2 Chron.", "Ezra", "Neh.", "Esther", "Job", "Ps.", "Prov.", "Eccl.", "Song of Sol.", "Isa.", "Jer.", "Lam.", "Ezek.", "Dan.", "Hos.", "Joel", "Amos", "Obad.", "Jonah", "Mic.", "Nah.", "Hab.", "Zeph.", "Hag.", "Zech.", "Mal.", "Matt.", "Mark", "Luke", "John", "Acts", "Rom.", "1 Cor.", "2 Cor.", "Gal.", "Eph.", "Phil.", "Col.", "1 Thess.", "2 Thess.", "1 Tim.", "2 Tim.", "Titus", "Philem.", "Heb.", "Jas.", "1 Pet.", "2 Pet.", "1 John", "2 John", "3 John", "Jude", "Rev."];
      break;

    case 'nwt_bookName_bce':
      return ["Genesis", "Exodus", "Leviticus", "Numbers", "Deuteronomy", "Joshua", "Judges", "Ruth", "1 Samuel", "2 Samuel", "1 Kings", "2 Kings", "1 Chronicles", "2 Chronicles", "Ezra", "Nehemiah", "Esther", "Job", "Psalm", "Proverbs", "Ecclesiastes", "Song of Solomon", "Isaiah", "Jeremiah", "Lamentations", "Ezekiel", "Daniel", "Hosea", "Joel", "Amos", "Obadiah", "Jonah", "Micah", "Nahum", "Habakkuk", "Zephaniah", "Haggai", "Zechariah", "Malachi"];
      break;

    case 'nwt_bookAbbrev1_bce':
      return ["Ge", "Ex", "Le", "Nu", "De", "Jos", "Jg", "Ru", "1Sa", "2Sa", "1Ki", "2Ki", "1Ch", "2Ch", "Ezr", "Ne", "Es", "Job", "Ps", "Pr", "Ec", "Ca", "Isa", "Jer", "La", "Eze", "Da", "Ho", "Joe", "Am", "Ob", "Jon", "Mic", "Na", "Hab", "Zep", "Hag", "Zec", "Mal"];
      break;

    case 'nwt_bookAbbrev2_bce':
      return ["Gen.", "Ex.", "Lev.", "Num.", "Deut.", "Josh.", "Judg.", "Ruth", "1 Sam.", "2 Sam.", "1 Ki.", "2 Ki.", "1 Chron.", "2 Chron.", "Ezra", "Neh.", "Esther", "Job", "Ps.", "Prov.", "Eccl.", "Song of Sol.", "Isa.", "Jer.", "Lam.", "Ezek.", "Dan.", "Hos.", "Joel", "Amos", "Obad.", "Jonah", "Mic.", "Nah.", "Hab.", "Zeph.", "Hag.", "Zech.", "Mal."];
      break;

    case 'nwt_bookName_ce':
      return ["Matthew", "Mark", "Luke", "John", "Acts", "Romans", "1 Corinthians", "2 Corinthians", "Galatians", "Ephesians", "Philippians", "Colossians", "1 Thessalonians", "2 Thessalonians", "1 Timothy", "2 Timothy", "Titus", "Philemon", "Hebrews", "James", "1 Peter", "2 Peter", "1 John", "2 John", "3 John", "Jude", "Revelation"];
      break;

    case 'nwt_bookAbbrev1_ce':
      return ["Mt", "Mr", "Lu", "Joh", "Ac", "Ro", "1Co", "2Co", "Ga", "Eph", "Php", "Col", "1Th", "2Th", "1Ti", "2Ti", "Tit", "Phm", "Heb", "Jas", "1Pe", "2Pe", "1Jo", "2Jo", "3Jo", "Jude", "Re"];
      break;

    case 'nwt_bookAbbrev2_ce':
      return ["Matt.", "Mark", "Luke", "John", "Acts", "Rom.", "1 Cor.", "2 Cor.", "Gal.", "Eph.", "Phil.", "Col.", "1 Thess.", "2 Thess.", "1 Tim.", "2 Tim.", "Titus", "Philem.", "Heb.", "Jas.", "1 Pet.", "2 Pet.", "1 John", "2 John", "3 John", "Jude", "Rev."];
      break;

    case 'single_chapter_bible_nums':
      return [31, 57, 63, 64, 65];
      break;

    // Make sure that entries below exist in function get_url > switch (bible_version)
    case 'bible_versions':
      return {
        nwt: 'New World Translation (NWT)',
        nwt_wol: 'New World Translation (NWT) (WOL)',
        nwtsty: 'NWT Study Bible',
        nwtsty_wol: 'NWT Study Bible (WOL)',
        nwtrbi8: 'NWT Reference Bible',
        nwtrbi8_wol: 'NWT Reference Bible (WOL)',
        kjv_jw: 'King James Version of 1611',
        kjw_wol: 'King James Version of 1611 (WOL)',
        by_jw: 'The Bible in Living English',
        by_wol: 'The Bible in Living English (WOL)',
        asv_jw: 'American Standard Version',
        asv_wol: 'American Standard Version (WOL)',
        ebr_jw: 'Emphasized Bible',
        ebr_wol: 'Emphasized Bible (WOL)'
      };
      break;

    default:
      undefined;
  }

}
