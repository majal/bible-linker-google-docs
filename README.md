# Bible linker for Google Docs
A Google Documents Apps Script and Add-on that searches for Bible verses and converts them to links. Choose from a selection of online Bible sources.

## How to use

### As an add-on
1. Install the **[Bible Linker add-on](https://workspace.google.com/marketplace/app/bible_linker/483504290926)** from Google Marketplace.  
For those who installed Bible Linker via Google Marketplace **before April 12, 2023,** you will need to reinstall the add-on as v2 required an additional [permission](https://developers.google.com/apps-script/add-ons/concepts/workspace-scopes) to make Bible Linker [modular](https://github.com/majal/bible-linker-google-docs/discussions/24#discussioncomment-5553877).
2. In your Google Document, find the add-on under **Extensions > Bible Linker.** (For those using Google Workspace accounts, the add-on may be found under Add-ons > Bible Linker.)

### As a stand-alone script 
1. In your Google Document, click **Extensions > Apps Script.**
2. In the page that will open, **copy and paste** the contents of **[Code.gs](Code.gs).**
3. **Save.** You may now close the Apps Script window and **refresh** your document.
4. Find Bible Linker in the **top-level Menu.** Enjoy!

**Tip:** If there is no selected text, Bible Linker will process the whole document. If you **highlighted a selection,** it will only create links in that selected portion of the document.

## Improving the script
Coders are welcome to improve the script. It's open source. MIT license.

### Things that could be worked on:
* Find and fix bugs
* ~~Support for other languages~~ Done in v2!
* ~~Support for more online Bibles~~ Done in v2!
* Check if Bible reference (chapter and verse) is valid
* Apply books exclusions in cases where some books are not available for the specific Bible version
* Add new languages by creating JSON files.  
For samples of valid JSON files, see the [bible-data](bible-data) directory.  
[Testing](https://www.google.com/search?q=json+validator) of new JSON files can be done via the menu **Extensions > Bible Linker > Choose language (data source) > Set custom data source ...**

## Other copyrights
The Bibles and languages initially supported came from [versions listed in jw.org](https://www.jw.org/en/library/bible/) and [Watchtower Online Library](https://wol.jw.org/en/wol/binav/r1/lp-e). The [Terms of Use](https://www.jw.org/finder?prefer=content&wtlocale=E&docid=1011511) of these websites do not allow the sharing of their contents. However, it allows [linking to these websites](https://www.jw.org/finder?prefer=content&wtlocale=E&docid=1011511&par=21-23), which is what this script does.
