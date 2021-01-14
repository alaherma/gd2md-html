/*
 * Copyright 2020 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

// *** OneCommunity.gs ***
// Google Doc to HTML (following this wed edition rules : https://docs.google.com/spreadsheets/d/1iEs9sBFSW1pAp40QhNAUy25Y3lxkDmQzxICAo8BiFtQ/edit?usp=sharing) conversion processing.
// May use some HTML conversions

var OneCommunity = OneCommunity || {
  // Attribute change markers.
  codeOpen:   '<code>',
  codeClose:  '</code>',
  italicOpen: '<em>',
  italicClose:'</em>',
  boldOpen:   '<strong>',
  boldClose:  '</strong>',
  
  // HTML code blocks (do not add \n at end of <pre>):
  openCodeBlock:         '\n\n<pre class="prettyprint">',
  openCodeBlockStart:    '\n\n<pre class="prettyprint lang-',
  openCodeBlockEnd:      '">',
  openCodeBlockLangNone: '\n\n<pre>',
  closeCodeBlock:        '</pre>\n\n',

  // non-semantic underline, since Docs supports it.
  underlineStart: '<span style="text-decoration:underline;">',
  underlineEnd:   '</span>'
};

OneCommunity.tablePrefix = '  ';

OneCommunity.doHtml = function(config) {
  // Basically, we can use the same code as doMarkdown, just change the markup.
  gdc.useHtml();
  
  gdc.config(config);
  // Get the body elements.
  var elements = gdc.getElements();

  // Main loop to walk through all the document's child elements.
  for (var i = 0, z = elements.length; i < z; i++) {
    OneCommunity.handleChildElement(elements[i]);
  }

  // Write image zip if selected (and available).
  if (gdc.zipImages) {
    if (izip) {
      izip.createImagesZip();
    }
  }
  
  if (gdc.hasImages) {
    gdc.info += '\n* This document has images: check for ' + gdc.alertPrefix;
    gdc.info += ' inline image link in generated source and store images to your server.';
    gdc.info += ' NOTE: Images in exported zip file from Google Docs may not appear in ';
    gdc.info += ' the same order as they do in your doc. Please check the images!\n';
  }
  
  if (gdc.hasFootnotes) {
    gdc.info += '\n* Footnote support in HTML is alpha: please check your footnotes.';
  }
  
  // Record elapsed time.
  var eTime = (new Date().getTime() - gdc.startTime)/1000;
  gdc.info = 'Conversion time: ' + eTime + ' seconds.\n' + gdc.info;
  // Add topComment (see gdc).
  gdc.info = gdc.topComment + gdc.info;

  // Warn at the top if DEBUG is true.
  if (DEBUG) {
    gdc.info = '<!-- WARNING: DEBUG is TRUE!! -->\n\n' + gdc.info;
  }

  // Add info and alert message to top of output.
  gdc.setAlertMessage();
  gdc.out = gdc.alertMessage + gdc.out;
  // Add info comment if desired.
  if (!gdc.suppressInfo) {
    gdc.out = gdc.info + '\n----->\n\n' + gdc.out;
  }

  var end = true;
  gdc.maybeCloseList(elements[i-1], end);

  // Output content.
  gdc.flushBuffer();
  gdc.flushFootnoteBuffer();

  // Close footnotes list if necessary.
  if (gdc.hasFootnotes) {
    gdc.writeStringToBuffer('\n</ol></div>');
    gdc.flushBuffer();
  }

  return gdc.out;
};

// Switch for handling different child elements for HTML conversion.
// Use for all element types, unless they have no children.
OneCommunity.handleChildElement = function(child) {
  gdc.useHtml();
  var childType = child.getType();
  
  // Get indent if possible for this element.
  // For HTML, we can also count blank paragraphs: Note difference in md.handleChildElement.
  if (child.getIndentStart) {
    gdc.indent = child.getIndentStart();
  }
  
  OneCommunity.checkList();

  // Most common element types first.
  switch (childType) {
    case PARAGRAPH:
      // Note that md.handleparagraph does both md and OneCommunity.
      md.handleParagraph(child);
      break;
    case TEXT:
      try {
        gdc.handleText(child);
      } catch(e) {
          gdc.log('ERROR handling text element:\n\n' + e + '\n\nText: ' + child.getText());
      }
      break;
    case LIST_ITEM:
      if (gdc.isTable || gdc.isHTML) {
        OneCommunity.handleListItem(child);
      } else {
        md.handleListItem(child);
      }
      break;
    case TABLE:
      OneCommunity.handleTable(child);
      break;
    case TABLE_ROW:
      OneCommunity.handleTableRow(child);
      break;
    case TABLE_CELL:
      OneCommunity.handleTableCell(child);
      break;
    case TABLE_OF_CONTENTS:
      gdc.isToc = true;
      gdc.handleTOC(child);
      gdc.isToc = false;
      break;
    case HORIZONTAL_RULE:
      break;
    case FOOTNOTE:
      OneCommunity.handleFootnote(child);
      break;
    case FOOTNOTE_SECTION:
      break;
    case FOOTER_SECTION:
      break;
    case INLINE_DRAWING:
      break;
    case INLINE_IMAGE:
      gdc.handleImage(child);
      break;
    case PAGE_BREAK:
      break;
    case EQUATION:
      break;
    case UNSUPPORTED:
      gdc.log('child element: UNSUPPORTED');
      break;
    default:
      gdc.log('child element: unknown');
  };
  gdc.lastChildType = childType;
};

OneCommunity.handleTable = function(tableElement) {
  // Note that we're converting all tables to HTML.
  if (!gdc.hasTables) {
    gdc.info += '\n* Tables are currently converted to HTML tables.';
    gdc.hasTables = true;
  }
 
  // init counters.
  gdc.nCols = 0;
  gdc.nRows = 0;

  // Check for a single-cell table first. Assume it's a code block. But add a warning to info. if long.
  // Guard against a spurious table with 0 rows.
  gdc.nRows = tableElement.getNumRows();
  if (gdc.nRows === 0) {
    return;
  }
  gdc.nCols = tableElement.getRow(0).getNumCells();

  // Handle single-cell table here.
  if (gdc.nCols === 1 && gdc.nRows === 1) {
    
    // Examine text length and text font to see if it's a suspicious single-cell table.
    var text = tableElement.getChild(0).getText();
    var textElement = tableElement.getChild(0).editAsText();
    // Warn if we find a really long single-cell table: might be an artifact of copying from Sites.
    // But if it's in code font (first and last lines), we'll let it go.
    // How long is suspiciously long?
    var singleCellMaxChars = 5120;
    if (text.length > singleCellMaxChars && !gdc.textIsCode(textElement) ) {
      gdc.warningCount++;
      gdc.info += '\nWARNING:\nFound a long single-cell table ';
      gdc.info += '(' + text.length + ' characters) starting with:\n';
      gdc.info += '**start sample:**\n';
      gdc.info += text.substring(0, 32) + '\n**end sample**\n';
      gdc.info += 'Check to make sure this is supposed to be a code block.\n';
      gdc.alert('Long single-cell table. Check to make sure this is meant to be a code block.');
    }
        
    // Markdown or HTML table for single-cell table.
    // Also handling lang for single-cell table.
    if (gdc.docType === gdc.docTypes.md && !gdc.isHTML) {
      gdc.inCodeBlock = gdc.isSingleCellTable = true;
      var text = tableElement.getText();
      var lang = gdc.getLang(text.split('\n')[0]);
      if (lang !== '') {
        // Skip first line, since it just specified lang.
        // XXX: note difference between HTML and Markdown. Why?
        text = text.substring(text.indexOf('\n') + 1);
      } 
      
      // Write the code block.
      gdc.startCodeBlock(lang);
      gdc.writeStringToBuffer(text);
      gdc.writeStringToBuffer('<newline>' + md.closeCodeBlock);
      gdc.inCodeBlock = gdc.isSingleCellTable = false;
    } else { // HTML code block.
      gdc.inCodeBlock = gdc.isSingleCellTable = true;
      var text = tableElement.getText();
      var lang = gdc.getLang(text.split('\n')[0]);
      if (lang !== '') {
        // Skip first line, since it just specified lang.
        // Note difference between HTML and Markdown. Why?
        text = text.substring(text.indexOf('\n'));
      }
      
      // Write the code block.
      gdc.startCodeBlock(lang);
      text = OneCommunity.escapeOpenTag(text);
      // Mark things that might get wrapped.
      text = util.markSpecial(text);
      text = util.markNewlines(text);
      gdc.writeStringToBuffer(text);
      gdc.writeStringToBuffer(OneCommunity.closeCodeBlock);      
      gdc.inCodeBlock = gdc.isSingleCellTable = false;
    }
    
    return;
  }
  
  // Regular table processing.
  gdc.isTable = true;
  
  gdc.useHtml();
  
  // Go through children of this table.
  gdc.writeStringToBuffer('\n\n<table>');
  md.childLoop(tableElement);
  gdc.writeStringToBuffer('\n</table>\n\n');
  // Turn off guard for table cell.
  gdc.startingTableCell = false;
  
  gdc.isTable = false;
  gdc.useMarkdown();
};

OneCommunity.handleTableRow = function(tableRowElement) {
  
  if (gdc.isSingleCellTable === true) {
    md.childLoop(tableRowElement);
    return;
  }
  
  // Go through children of this row.
  gdc.writeStringToBuffer('\n  <tr>');
  md.childLoop(tableRowElement);
  gdc.writeStringToBuffer('\n  </tr>');
};

OneCommunity.handleTableCell = function(tableCellElement) {
  if (gdc.isSingleCellTable === true) {
    md.childLoop(tableCellElement);
    return;
  }

  gdc.startingTableCell = true;

  // Set <td> attribute for colspan or rowspan, if necessary (>1).
  var tdAttr = '';
  
  // Rowspan handling.
  var rowspan = tableCellElement.getRowSpan();
  if (rowspan === 0) { return; }

  if (rowspan > 1) {
    tdAttr += ' rowspan="' + rowspan + '"';
  }
  // End rowspan code.

  // Colspan handling.
  var colspan = tableCellElement.getColSpan();
  // Skip cells that have been merged over. 
  if (colspan === 0) { return; }
    
  if (colspan > 1) {
    tdAttr += ' colspan="' + colspan + '"';
  }
  // End colspan code.
  
  // Add attribute only if non-empty.
  if (tdAttr) {
    gdc.writeStringToBuffer('\n   <td' + tdAttr + ' >');
  } else {
    gdc.writeStringToBuffer('\n   <td>');
  }
  
  // Go through children of this cell.
  md.childLoop(tableCellElement);
  md.maybeEndCodeBlock();
  OneCommunity.closeAllLists();
  gdc.writeStringToBuffer('\n   </td>');
};

// Handle the heading type of the paragraph.
// Need to close heading for HTML too, so need to save heading state.
// Fall through for NORMAL.
OneCommunity.handleHeading = function(heading, para) {
  
  // We're doing a little dance here for heading demotion. Also for closing tags.
  var htitle = 0;
  if (gdc.demoteHeadings) {
    htitle = 1;
  }
  switch (heading) {
    case DocumentApp.ParagraphHeading.HEADING6:     
     // Warn about level 6 headings if demoting (and do not demote h6).
      var warning = 'H6 not demoted to H7.';
      if (gdc.demoteHeadings) {
        if (!gdc.warnedAboutH7) {
          gdc.warn(warning + ' Look for "' + warning + '" inline.');
          gdc.warnedAboutH7 = true;
        }
        gdc.writeStringToBuffer('\n<!--' + warning + ' -->\n');
      }
      gdc.writeStringToBuffer('\n<h6');
      OneCommunity.isHeading = OneCommunity.h6 = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING5:
      gdc.writeStringToBuffer('\n<h' + (5+htitle) );
      OneCommunity.isHeading = OneCommunity.h5 = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING4:
      gdc.writeStringToBuffer('\n<h' + (4+htitle) );
      OneCommunity.isHeading = OneCommunity.h4 = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING3:
      gdc.writeStringToBuffer('\n<h' + (3+htitle) );
      OneCommunity.isHeading = OneCommunity.h3 = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING2:
      gdc.writeStringToBuffer('\n<h' + (2+htitle) );
      OneCommunity.isHeading = OneCommunity.h2 = true;
      break;
    case DocumentApp.ParagraphHeading.HEADING1:
      if (!gdc.demoteHeadings) {
        gdc.h1Count++;
      }
    case DocumentApp.ParagraphHeading.TITLE:
      gdc.writeStringToBuffer('\n<h' + (1+htitle) );
      OneCommunity.isHeading = OneCommunity.h1 = true;
      break;
    
    // Handle SUBTITLE as a regular paragraph.
    case DocumentApp.ParagraphHeading.SUBTITLE:
    default:
      OneCommunity.isHeading = false;
      // Add paragraph markup if appropriate.
      if (!gdc.isMarkdown && !gdc.inCodeBlock && !gdc.isTable) {
        gdc.writeStringToBuffer('\n<p');
      }
  }
  
  // Insert id for HTML headings (that occur after the TOC).
  // But do not warn unless we're actually linking to a heading that has no ID.
  var id = gdc.headingIds[para.getText()];
  if (id) {
    gdc.writeStringToBuffer(' id="' + gdc.headingIds[para.getText()] + '"');
  }
  
  // Close the tag.
  gdc.writeStringToBuffer('>');
};

// Close heading if necessary. (Blank line after to keep Markdown parser happy.)
OneCommunity.closeHeading = function() {

  // We're doing a little dance here for heading demotion.
  var htitle = 0;
  if (gdc.demoteHeadings) {
    htitle = 1;
  }
  if (OneCommunity.h1)      { OneCommunity.h1 = false; gdc.writeStringToBuffer('</h' + (1+htitle) + '>\n\n'); }
  else if (OneCommunity.h2) { OneCommunity.h2 = false; gdc.writeStringToBuffer('</h' + (2+htitle) + '>\n\n'); }
  else if (OneCommunity.h3) { OneCommunity.h3 = false; gdc.writeStringToBuffer('</h' + (3+htitle) + '>\n\n'); }
  else if (OneCommunity.h4) { OneCommunity.h4 = false; gdc.writeStringToBuffer('</h' + (4+htitle) + '>\n\n'); }
  else if (OneCommunity.h5) { OneCommunity.h5 = false; gdc.writeStringToBuffer('</h' + (5+htitle) + '>\n\n'); }
  else if (OneCommunity.h6) { OneCommunity.h6 = false; gdc.writeStringToBuffer('</h' + (6+htitle) + '>\n\n'); }

  OneCommunity.isHeading = false;
};

// Formats footnotes for HTML. For HTML, we'll print out the actual
// footnotes at the end.
OneCommunity.handleFootnote = function(footnote) {

  gdc.hasFootnotes = true;
  
  gdc.footnoteNumber++;
  var fSection = footnote.getFootnoteContents();
  if (!fSection) {
    // Index is one less.
    var findex = gdc.footnoteNumber - 1;
    fSection = gdc.footnotes[findex].getFootnoteContents();
  }
  
  // Write the footnote ref link in the text.
  gdc.writeStringToBuffer('<sup id="fnref' + gdc.footnoteNumber + '"><a href="#fn' 
    + gdc.footnoteNumber + '" rel="footnote">' + gdc.footnoteNumber + '</a></sup>');

  // Now, write the footnotes themselves.
  gdc.isFootnote = true;
  // Open list for first footnote.
  if (gdc.footnoteNumber === 1) {
      gdc.writeStringToBuffer('\n\n<!-- Footnotes themselves at the bottom. -->'
      + '\n\n<h2>Notes</h2>'
      + '\n<div class="footnotes">'
      + '\n<hr>'
      + '\n<ol>');
  }
  // Each HTML footnote is a list item in an ordered list:
  gdc.writeStringToBuffer('<li id="fn' + gdc.footnoteNumber + '">');
  md.childLoop(fSection);
  // Close footnote with a link back to the ref.
  gdc.writeStringToBuffer('&nbsp;<a href="#fnref' + gdc.footnoteNumber + '" rev="footnote">&#8617;</a>');
  gdc.isFootnote = false;
};

// HTML list stuff.
// For keeping track of nested HTML lists.
OneCommunity.listStack = [];

OneCommunity.handleListItem = function(listItem) {
  
  // Close definition list if we're in one now.
  if (gdc.inDlist) {
    gdc.closeDlist();
  }
  
  var gt = listItem.getGlyphType(),
      textElement = listItem.asText(),
      attrix = textElement.getTextAttributeIndices(),
      isList = true;

  OneCommunity.nestingLevel = listItem.getNestingLevel();

  // Check if we're in a code block and end if so.
  if (gdc.inCodeBlock) {
    gdc.writeStringToBuffer(OneCommunity.closeCodeBlock);
    gdc.inCodeBlock = false;
  }


  gdc.listPrefix = '';
  for (var i = 0; i < OneCommunity.nestingLevel; i++) {
    gdc.listPrefix += ' ';
  }
  
  // Determine what type of list before we open it (if necessary).
  if (gt === DocumentApp.GlyphType.BULLET
      || gt === DocumentApp.GlyphType.HOLLOW_BULLET
      || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
    gdc.listType = gdc.ul;
    OneCommunity.maybeOpenList(listItem);
  } else {
    gdc.listType = gdc.ol;
    OneCommunity.maybeOpenList(listItem);
  }
  
  gdc.writeStringToBuffer('\n');
  // Note that ulItem, olItem are the same in HTML (<li>).
  gdc.writeStringToBuffer(gdc.listPrefix + gdc.htmlMarkup.ulItem);
  md.childLoop(listItem);
  
  // Check to see if we should close this list.
  gdc.maybeCloseList(listItem);
};

// Called to check if we're exiting a list as we enter a new element.
// Maybe this should just be gdc.checklist().
OneCommunity.checkList = function() {
  if (gdc.isList && !gdc.indent) {
    OneCommunity.closeAllLists();
    gdc.isList = false;
  }
};
// Closes list item. Not necessary for Markdown.
OneCommunity.closeListItem = function() {
  gdc.writeStringToBuffer(gdc.markup.liClose);
};
OneCommunity.maybeOpenList = function (listItem) {
  // Do we need to open a list?
  var previous = listItem.getPreviousSibling();
  var previousType;
  if (previous) {
    previousType = previous.getType();
  }
  // Open list if last sibling was not a list item.
  if (previousType !== DocumentApp.ElementType.LIST_ITEM) {
    OneCommunity.openList();
  } else
    // Open a new list if nesting level increases.
    if (OneCommunity.nestingLevel > previous.getNestingLevel()) {
    OneCommunity.openList();
  }
};
// Open list and save current list type to stack.
OneCommunity.openList = function() {
  gdc.isList = true;
  if (gdc.nestingLevel === 0) {
    gdc.writeStringToBuffer('\n');
  }

  if (gdc.ul === gdc.listType) {
    gdc.writeStringToBuffer(gdc.listPrefix + gdc.htmlMarkup.ulOpen);
    // Add to front of list stack.
    OneCommunity.listStack.unshift(gdc.ul);
  }
  if (gdc.ol === gdc.listType) {
    gdc.writeStringToBuffer(gdc.listPrefix + gdc.htmlMarkup.olOpen);
    OneCommunity.listStack.unshift(gdc.ol);
  }
};
// Close list and remove it's list type from the stack.
OneCommunity.closeList = function() {
  // Close the last item of the list.
  OneCommunity.closeListItem();
  if (OneCommunity.listStack[0] === gdc.ul) {
    gdc.writeStringToBuffer(gdc.listPrefix + gdc.htmlMarkup.ulClose);
  }
  if (OneCommunity.listStack[0] === gdc.ol) {
    gdc.writeStringToBuffer(gdc.listPrefix + gdc.htmlMarkup.olClose);
  }
  OneCommunity.listStack.shift();
  if (OneCommunity.listStack.length === 0) {
    gdc.isList = false;
  }
};
// But what about a table that's in a list item?
OneCommunity.closeAllLists = function() {
  var list = OneCommunity.listStack[0];
  while (OneCommunity.listStack.length > 0) {
    var list = OneCommunity.listStack[0];
    OneCommunity.closeList();
  }
};

// Escape angle brackets so code blocks will display HTML tags.
OneCommunity.escapeOpenTag = function(text) {
  text = text.replace(/</g, '&lt;');
  return text;
};
