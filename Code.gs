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
  if (last_used_bible_version == null) last_used_bible_version = 'wol_nwtsty';

  // Pull supported Bible versions
  const bible_versions = consts('bible_versions');
  const bible_versions_keys = Object.keys(bible_versions);

  // Set lastest used Bible version to the menu
  let last_used_bible_label = bible_versions[last_used_bible_version];
  let function_name = 'bible_linker_' + last_used_bible_version;
  let search_label = '🔗⠀Link verses using ' + last_used_bible_label;

  // Set menu
  var ui = DocumentApp.getUi();
  var menu = ui.createMenu('Bible Linker').addItem(search_label, function_name).addSeparator();
  var submenu_bible_ver = ui.createMenu('📖⠀Select Bible version');

  // Set dynamic submenus
  for (let n=0; n < bible_versions_keys.length; n++) {
    let key = bible_versions_keys[n];
    let function_name = 'bible_linker_' + key;
    let last_used_pointer = (last_used_bible_version == key) ? '▸ ⠀' : '⠀⠀';
    submenu_bible_ver.addItem(last_used_pointer + bible_versions[key], function_name);
  }

  // Create menu 
  menu.addSubMenu(submenu_bible_ver).addToUi();

}


// Dynamic menu hack
const bible_versions = consts('bible_versions');
const bible_versions_keys = Object.keys(bible_versions);
for (let n=0; n < bible_versions_keys.length; n++) {
  let key = bible_versions_keys[n];
  let function_name = 'bible_linker_' + key;
  this[function_name] = function() { bible_linker(key); };
}


function bible_linker(bible_version) {

  // Set the latest used Bible version
  if (bible_version == undefined || bible_version == null) bible_version = 'wol_nwtsty';
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('last_used_bible_version', bible_version);
  create_menu();

  // Get names of Bible books
  var nwt_bookName = consts('nwt_bookName');
  var nwt_bookAbbrev1 = consts('nwt_bookAbbrev1');
  var nwt_bookAbbrev2 = consts('nwt_bookAbbrev2');

  // Run parser for each Bible name
  for (let n=0; n < nwt_bookName.length; n++) {
    parse_scripture(bible_version, nwt_bookName[n], n+1);
  }
  for (let n=0; n < nwt_bookAbbrev1.length; n++) {
    parse_scripture(bible_version, nwt_bookAbbrev1[n], n+1);
  }
  for (let n=0; n < nwt_bookAbbrev2.length; n++) {
    parse_scripture(bible_version, nwt_bookAbbrev2[n], n+1);
  }

}


function parse_scripture(bible_version, bible_name, bible_num) {

  // Initialize Google Docs
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Variable(s) and constant(s)
  var single_chapters = consts('single_chapter_bible_nums');

  // Search for Bible references
  if (single_chapters.includes(bible_num)) {
    var ref_search_string = bible_name + ' [0-9 ,-]+';
  } else {
    var ref_search_string = bible_name + ' [0-9]+:[0-9 ,;:-]+';
  }
  var bibleref_full = body.findText(ref_search_string);
  
  // Cycle through each Bible reference found
  while (bibleref_full != null) {

    // Set reference start and end
    var bibleref_full_start = bibleref_full.getStartOffset();
    var bibleref_full_end = bibleref_full.getEndOffsetInclusive();

    // Isolate reference only
    var bibleref_full_astext = bibleref_full.getElement().asText();
    var bibleref_full_text = bibleref_full_astext.getText();
    var bibleref_full_text_slice = bibleref_full_text.slice(bibleref_full_start, bibleref_full_end + 1);

    // Split at semicolon (;)
    var bibleref_split = bibleref_full_text_slice.split(';');
  
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
          select_start = bibleref_full_start + offset_reference;
          select_end = select_start + url_text_len - 1;
          
          // Set links
          bibleref_full_astext.setLinkUrl(select_start, select_end, url);
          
          // Add to reference offset
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
        select_start = bibleref_full_start + offset_reference;
        select_end = select_start + url_text_len - 1;
        
        // Set links
        bibleref_full_astext.setLinkUrl(select_start, select_end, url);

        // Add to reference offset
        offset_reference += url_text_len + 2
      }
    }

    // Find the next match
    bibleref_full = body.findText(ref_search_string, bibleref_full);
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
  
    case 'wol_nwtsty':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/nwtsty/';
      url_hash = '#study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;
      
    case 'wol_nwt':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/nwt/';
      url_hash = '#study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
      }
    break;

    case 'wol_nwtrbi8':
      url_head = 'https://wol.jw.org/en/wol/b/r1/lp-e/Rbi8/';
      url_hash = '#study=discover&v=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + '/' + chapter_start + url_hash + bible_name_num + ':' + chapter_start + ':' + verse_start;
      } else {
        return url_head + bible_name_num + '/' + chapter_start + url_hash + bible_name_num + ':' + chapter_start + ':' + verse_start + '-' + bible_name_num + ':' + chapter_end + ':' + verse_end;
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

    case 'nwt':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=nwt&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
      }
    break;

    case 'rbi8':
      url_head = 'https://www.jw.org/finder?wtlocale=E&pub=Rbi8&bible=';

      if (chapter_start == chapter_end && verse_start == verse_end) {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0');
      } else {
        return url_head + bible_name_num + String(chapter_start).padStart(3, '0') + String(verse_start).padStart(3, '0') + '-' + bible_name_num + String(chapter_end).padStart(3, '0') + String(verse_end).padStart(3, '0');
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

    case 'single_chapter_bible_nums':
      return [31, 57, 63, 64, 65];
      break;

    // Make sure that entries below exist in function get_url > switch (bible_version)
    case 'bible_versions':
      return {
        wol_nwt: 'WOL - New World Translation (NWT)',
        wol_nwtsty: 'WOL - NWT Study Bible',
        wol_nwtrbi8: 'WOL - NWT Reference Bible',
        nwt: 'New World Translation (NWT)',
        nwtsty: 'NWT Study Bible',
        rbi8: 'NWT Reference Bible'
      };
      break;

    default:
      undefined;
  }

}
