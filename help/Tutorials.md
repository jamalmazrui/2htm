# Tutorials

## Contents

- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Read a Word Report in the Browser](#3-read-a-word-report-in-the-browser)
- [4 - Make a PDF Readable](#4-make-a-pdf-readable)
- [5 - Make Plain Text to Listen To](#5-make-plain-text-to-listen-to)
- [6 - Convert Slides and Spreadsheets](#6-convert-slides-and-spreadsheets)
- [7 - Convert a Whole Folder](#7-convert-a-whole-folder)
- [8 - Convert from Explorer and the Command Line](#8-convert-from-explorer-and-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Read a Word Report in the Browser](#3-read-a-word-report-in-the-browser)
- [4 - Make a PDF Readable](#4-make-a-pdf-readable)
- [5 - Make Plain Text to Listen To](#5-make-plain-text-to-listen-to)
- [6 - Convert Slides and Spreadsheets](#6-convert-slides-and-spreadsheets)
- [7 - Convert a Whole Folder](#7-convert-a-whole-folder)
- [8 - Convert from Explorer and the Command Line](#8-convert-from-explorer-and-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [0 - Overview](#0-overview)
- [1 - User Interface](#1-user-interface)
- [2 - Install and Launch](#2-install-and-launch)
- [3 - Read a Word Report in the Browser](#3-read-a-word-report-in-the-browser)
- [4 - Make a PDF Readable](#4-make-a-pdf-readable)
- [5 - Make Plain Text to Listen To](#5-make-plain-text-to-listen-to)
- [6 - Convert Slides and Spreadsheets](#6-convert-slides-and-spreadsheets)
- [7 - Convert a Whole Folder](#7-convert-a-whole-folder)
- [8 - Convert from Explorer and the Command Line](#8-convert-from-explorer-and-the-command-line)
- [9 - Conclusion](#9-conclusion)
- [1. Where to Go Next](#1-where-to-go-next)

<!-- walkthrough: written by makeTutorials.py, do not edit between the markers -->

## 0 - Overview

What 2htm is, the ways to run it, and the ten walks that teach it.

**Before you start:** Nothing to set up; this walk only listens.

### Step 1

Welcome. This is the first of ten short walks through 2htm. I am the host; the other voice is the screen reader, speaking as it would on your own computer.

### Step 2: Insert+Up Arrow

Two reader keys help in every walk. If a line goes by too fast, Insert plus Up Arrow says it again.

Screen reader:

- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Insert+Tab

And whenever you are unsure where you are, Insert plus Tab says what has the focus: its name, its kind and its value.

Screen reader:

- Source files edit, blank

Confirm the wording with a live run: buildTutorials -live.

### Step 4

2htm converts documents into web pages that read well with a screen reader: Word documents, PDF files, PowerPoint slides, Excel workbooks and Markdown.

### Step 5

A web page keeps the document's headings, lists and tables, so you move through it with the same keys as any web page: H for the next heading, T for a table, L for a list.

### Step 6

It can make plain text instead, for listening with a speech synthesizer or pasting into an email, with the pictures left out.

### Step 7

There are three ways to run it: a small dialog opened from the desktop, an entry in File Explorer's context menu, and the command line for many files at once.

### Step 8

Here are the walks. One, the dialog. Two, installing. Three, a Word report in the browser. Four, a PDF made readable. Five, plain text to listen to.

### Step 9

Six, slides and spreadsheets. Seven, a whole folder at once. Eight, File Explorer and the command line. Nine, the conclusion, a glossary, and where to find help.

### Step 10

Each walk is a few minutes long and builds on the ones before it, so they are best heard in order the first time. Walk one comes next.

**Something to try:** Think of one document you would rather read in a browser. Walks three to six convert one like it.

## 1 - User Interface

The 2htm dialog: its fields, its underlined letters, and the keys that work anywhere in it, each shown as the reader speaks it.

**Before you start:** 2htm installed; the dialog open with Alt+Control+2.

### Step 1

One want: to move around the 2htm dialog without hunting. Everything in it can be reached by a letter.

### Step 2: Alt+Control+2

I press the desktop key for 2htm. The dialog opens on its first field.

Screen reader:

- 2htm dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3

Source files names what to convert: one file, a folder, a pattern such as star dot docx, or several with spaces between. Its letter is S.

### Step 4: Tab

Tab moves to the next control, a button that opens a file picker.

Screen reader:

- Browse source button

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Alt+O

Every label has an underlined letter, and Alt with that letter jumps straight to it. Output directory is O.

Screen reader:

- Output directory edit

Confirm the wording with a live run: buildTutorials -live.

### Step 6

Left blank, each web page is written beside its document. Choose output, C, picks a folder instead of typing one.

### Step 7: Alt+F

Then the check boxes. Force replacements, F, replaces a web page already made; without it, a document whose page exists is skipped.

Screen reader:

- Force replacements check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Alt+P

Plain text, P, makes a text file instead of a web page.

Screen reader:

- Plain text check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Alt+I

Strip images, I, leaves out the pictures.

Screen reader:

- Strip images check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 10

View output, V, opens the folder when the run is done; Log session, L, writes a log beside the output; Use configuration, U, remembers these choices for next time.

### Step 11: Space

Space checks a box and Space again clears it.

Screen reader:

- checked

Confirm the wording with a live run: buildTutorials -live.

### Step 12: Shift+F1

Not sure what something does? Shift plus F1 says the tip for the control you are on.

Screen reader:

- Strip images: leave pictures out of the output

Confirm the wording with a live run: buildTutorials -live.

### Step 13: F1

F1 shows Help: every control with its tip, and every key.

Screen reader:

- 2htm Help dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 14: Escape

Escape closes Help and returns to the same place.

Screen reader:

- Strip images check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 15: Space

Space clears the box again, as I found it.

Screen reader:

- not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 16: F7

F7 lists every control in the dialog; choose one and press Enter to move there.

Screen reader:

- Controls list

Confirm the wording with a live run: buildTutorials -live.

### Step 17: Enter

I choose Source files.

Screen reader:

- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 18: Tab

Enter starts converting from any field, and Control plus Enter does the same from any control. Escape closes the dialog without converting.

Screen reader:

- Browse source button

Confirm the wording with a live run: buildTutorials -live.

### Step 19: Alt+G

Two more buttons sit at the end. Guide, G, opens the full guide in the browser and returns here; Default settings, D, clears every field and box.

Screen reader:

- Guide button

Confirm the wording with a live run: buildTutorials -live.

### Step 20: Up Arrow

The text fields remember your answers. In Source files, Up Arrow brings back what you typed last time, and earlier answers above it.

Screen reader:

- C colon backslash Documents backslash Annual Report dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 21: Down Arrow

Down Arrow returns toward an empty field, so a fresh entry is a key or two away.

Screen reader:

- Source files edit, blank

Confirm the wording with a live run: buildTutorials -live.

### Step 22

Every label you hear is the control's own name, read once by the screen reader; 2htm adds speech only for what the reader cannot know, such as progress during a long run.

### Step 23

That is the whole dialog: what to convert, where it goes, five boxes, and a key for everything. The next walk installs 2htm.

**Something to try:** Open the dialog, press F7, choose a control from the list, then press Shift+F1 there to hear its tip.

## 2 - Install and Launch

Installing 2htm for everyone on the computer, the desktop key that opens it, and the File Explorer entry that converts one file.

**Before you start:** 2htm_setup.exe downloaded from the 2htm releases page on GitHub.

### Step 1

One want: 2htm on this computer, ready from the desktop and from File Explorer, in one sitting.

### Step 2: Enter

I open the downloaded setup program. It installs for everyone on the computer, so Windows asks for administrator rights; I say yes.

Screen reader:

- Setup - 2htm dialog
- Welcome to the 2htm Setup Wizard

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Enter

The welcome page summarizes the license: 2htm is free and open source, under the MIT license. Enter takes Next.

Screen reader:

- Select Destination Location

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Enter

The first time, it asks for a folder; the default in Program Files is right. An update goes to the same place without asking.

Screen reader:

- Ready to Install

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Enter

Enter again, and it copies the program and its documents.

Screen reader:

- Installing

Confirm the wording with a live run: buildTutorials -live.

### Step 6

The last page offers to launch 2htm now, already checked, and to open the user guide.

Screen reader:

- Launch 2htm now check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Enter

I press Finish. A short box says what was installed and where the logs are.

Screen reader:

- 2htm setup complete

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Enter

When I close that box, 2htm starts, ready on Source files.

Screen reader:

- 2htm dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 9

Setup added three ways to start. The first is the desktop shortcut whose key is Alt plus Control plus 2, which opens the dialog from anywhere.

### Step 10

The second is the Start menu, where 2htm has a group with the program, the guide and the uninstaller.

### Step 11: Escape

The third is in File Explorer. I close the dialog and go to a Word document there.

Screen reader:

- Desktop

Confirm the wording with a live run: buildTutorials -live.

### Step 12: Shift+F10

Shift plus F10 opens the context menu for the file I am on.

Screen reader:

- Context menu

Confirm the wording with a live run: buildTutorials -live.

### Step 13: 2

The number 2 is Convert via 2htm. It converts this one file at once, and the web page lands beside it.

Screen reader:

- 2htm: 1 file converted

Confirm the wording with a live run: buildTutorials -live.

### Step 14

Office must be installed to convert Word, PowerPoint and Excel files, since 2htm asks Office to read them. PDF and Markdown files need nothing else.

### Step 15: F11

Keeping 2htm current takes one key: in the dialog, F11 asks GitHub whether a newer version is out, and Enter installs it when there is.

Screen reader:

- Checking for a newer version

Confirm the wording with a live run: buildTutorials -live.

### Step 16: Enter

If this copy is the newest, No is the default, and Enter returns to the dialog.

Screen reader:

- 2htm dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 17

Your settings and logs live in your own folders, so an update never touches them, and uninstalling leaves your converted pages where you saved them.

### Step 18

The program itself sits in the exec folder of the installation, beside its documents in the help folder, so the guide is there even without the dialog open.

### Step 19

Each run writes a session log in your own application data, kept whatever the dialog says, so a problem can always be traced afterwards.

### Step 20

If setup ever stops partway, its own log says why, and running the setup again finishes the job.

### Step 21

To remove 2htm, use Installed apps in Windows Settings, or the uninstaller in its Start menu group.

Screen reader:

- 2htm, Uninstall button

Confirm the wording with a live run: buildTutorials -live.

### Step 22

Your converted pages are never removed with it, since they are wherever you saved them.

### Step 23

Installed, launched three ways, and kept current with F11. The next walk converts a real document.

**Something to try:** After installing, right-click a document in File Explorer and find Convert via 2htm.

## 3 - Read a Word Report in the Browser

Converting a long Word report into a web page, then reading it by headings, lists and tables in the browser.

**Before you start:** A long Word document of your own, and Microsoft Word installed.

### Step 1

One want: a forty-page Word report that I can skim by heading, the way I skim a web page, without Word's own interface in the way.

### Step 2: Alt+Control+2

I open the dialog. The focus is on Source files.

Screen reader:

- 2htm dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Alt+B

Browse source opens a file picker, rather than typing the path.

Screen reader:

- Open dialog
- File name edit

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Enter

I type the start of the report's name and press Enter. Its full path is now in Source files.

Screen reader:

- Source files edit
- C colon backslash Documents backslash Annual Report dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Alt+V

I tick View output, so the folder opens when the page is ready.

Screen reader:

- View output check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 6: Enter

Enter converts. Word reads the document in the background, so nothing appears on screen.

Screen reader:

- Converting Annual Report dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 7

A short box reports the result.

Screen reader:

- 1 file converted
- OK button

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Enter

Enter closes it, and File Explorer opens on the folder, on the new web page, Annual Report dot h t m.

Screen reader:

- Annual Report dot h t m

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Enter

Enter opens it in the browser, where the screen reader is in its browse mode, ready to read.

Screen reader:

- Annual Report, heading level 1

Confirm the wording with a live run: buildTutorials -live.

### Step 10: H

H moves to the next heading, so the report's outline is heard one section at a time.

Screen reader:

- Summary, heading level 2

Confirm the wording with a live run: buildTutorials -live.

### Step 11: H

H again.

Screen reader:

- Findings, heading level 2

Confirm the wording with a live run: buildTutorials -live.

### Step 12: Insert+F6

Insert plus F6 lists every heading, the quickest way to jump to the section I want.

Screen reader:

- Headings list

Confirm the wording with a live run: buildTutorials -live.

### Step 13: T

Escape returns to the page. T moves to the next table, which keeps its rows and columns, so table keys read each cell with its header.

Screen reader:

- table with 4 columns and 6 rows

Confirm the wording with a live run: buildTutorials -live.

### Step 14: L

Lists keep their shape too: L moves to the next list and the reader says how many items it holds.

Screen reader:

- list with 5 items

Confirm the wording with a live run: buildTutorials -live.

### Step 15: G

And pictures keep the descriptions written in Word, so a chart says what the author wrote for it.

Screen reader:

- graphic, Sales by quarter, rising each quarter

Confirm the wording with a live run: buildTutorials -live.

### Step 16

The page is one file, with its look built in, so it can be emailed, kept or opened offline without anything else.

### Step 17: Alt+F

Converting again later, after the report changes, needs Force replacements, or 2htm keeps the page it already made.

Screen reader:

- Force replacements check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 18: K

Links in the report stay links. K moves to the next one, and Enter follows it, whether it points into the report or out to the web.

Screen reader:

- link, Appendix A

Confirm the wording with a live run: buildTutorials -live.

### Step 19: Enter

A link to a heading inside the report jumps there, which turns a long table of contents into a real one.

Screen reader:

- Appendix A, heading level 2

Confirm the wording with a live run: buildTutorials -live.

### Step 20: Alt+Left Arrow

Alt plus Left Arrow comes back, as in any browser.

Screen reader:

- link, Appendix A

Confirm the wording with a live run: buildTutorials -live.

### Step 21

The page's title, read when it opens and in the window's title bar, is the document's own title, so a list of open pages says what each one is.

### Step 22

Footnotes become a list at the end, each linked from its number in the text and back again.

### Step 23

Converting a report before a meeting means reading it at your own pace, by structure, rather than page by page in Word.

### Step 24

A long report, read like a web page, by headings, tables and lists. The next walk makes a PDF readable.

**Something to try:** Convert a report of your own, then press H in the browser to hear its headings in order.

## 4 - Make a PDF Readable

Converting a PDF into a web page whose text can be read, searched and copied, and what to expect from a scanned one.

**Before you start:** A PDF of your own, such as a statement or a manual.

### Step 1

One want: a PDF manual whose text I can read line by line and search, instead of a viewer that reads it in pieces.

### Step 2: Alt+Control+2

I open the dialog and type the PDF's path in Source files. Typing works as well as browsing.

Screen reader:

- 2htm dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Enter

Enter converts. No other program is needed for a PDF; 2htm reads it itself.

Screen reader:

- Converting Manual dot pdf

Confirm the wording with a live run: buildTutorials -live.

### Step 4

The result box confirms it.

Screen reader:

- 1 file converted

Confirm the wording with a live run: buildTutorials -live.

### Step 5

I open the new page in the browser. The text arrives as paragraphs, in reading order.

Screen reader:

- Manual, heading level 1

Confirm the wording with a live run: buildTutorials -live.

### Step 6

A PDF does not always mark its headings the way a Word document does, so a page may have fewer headings; paragraphs and lines are still there to read.

### Step 7: Control+F

Control plus F, the browser's find, searches the whole manual at once.

Screen reader:

- Find on page, edit

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Enter

I type warranty and press Enter. The browser lands on the first match.

Screen reader:

- warranty, 1 of 3

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Escape

Escape closes the find box, and the reading cursor is at that place.

Screen reader:

- The warranty covers parts and labor for two years.

Confirm the wording with a live run: buildTutorials -live.

### Step 10: Shift+Down Arrow

Text can be selected and copied as on any web page, which a PDF viewer often makes hard: Shift with Down Arrow selects, Control plus C copies.

Screen reader:

- The warranty covers parts and labor for two years. selected

Confirm the wording with a live run: buildTutorials -live.

### Step 11

A planned misstep: a scanned PDF, which is a picture of each page with no text inside it.

Screen reader:

- Converting Scan dot pdf

Confirm the wording with a live run: buildTutorials -live.

### Step 12

The page that results is nearly empty, because there was no text to take.

Screen reader:

- Scan, heading level 1
- blank

Confirm the wording with a live run: buildTutorials -live.

### Step 13

A scanned page needs text recognition first. HomerScribe and other tools read such pages; then 2htm can convert the text they find.

### Step 14

A quick test tells the two apart: if the browser's find cannot find a word you can see in the PDF, the PDF is a scan.

### Step 15: Control+End

A long PDF becomes one long page. Control plus End jumps to its end, and Control plus Home back to the top.

Screen reader:

- End of page

Confirm the wording with a live run: buildTutorials -live.

### Step 16

The browser remembers your place while the page stays open, so a manual can be read in sittings.

### Step 17

To keep one section, select it from its heading to the next, and paste it into a document of its own.

### Step 18

Page numbers printed in the PDF arrive as ordinary lines of text, so the page of a passage can still be found when someone asks for it.

### Step 19

Tables in a PDF are drawn rather than described, so they may arrive as lines of text rather than a table; the words are all there.

### Step 20

When a PDF must stay a PDF, the converted page is still a quick way to read and search it first.

### Step 21

For a PDF with text, one Enter makes it readable and searchable. The next walk makes plain text to listen to.

**Something to try:** Convert a PDF, then search the web page for one word with Control+F.

## 5 - Make Plain Text to Listen To

Converting a document to plain text, with the pictures left out, for a speech synthesizer, a braille display, or an email.

**Before you start:** A document of your own; any of the supported kinds.

### Step 1

One want: a document as clean text, to listen to in a synthesizer program on a walk, or to paste into an email without formatting.

### Step 2: Alt+Control+2

I open the dialog and put the document's path in Source files.

Screen reader:

- 2htm dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Alt+P

Plain text, P, makes a text file instead of a web page.

Screen reader:

- Plain text check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Space

Space checks it.

Screen reader:

- checked

Confirm the wording with a live run: buildTutorials -live.

### Step 5: Alt+I

Strip images, I, leaves out the lines that stood for pictures, which read as noise when listened to.

Screen reader:

- Strip images check box, not checked

Confirm the wording with a live run: buildTutorials -live.

### Step 6: Space

Space checks it too.

Screen reader:

- checked

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Enter

Enter converts. The result is a text file with the same name, ending in dot t x t.

Screen reader:

- 1 file converted

Confirm the wording with a live run: buildTutorials -live.

### Step 8

Plain text keeps the words and their order, paragraph by paragraph; headings become lines of their own, and tables become rows of text.

### Step 9

I open the text file in Notepad. The first line is the title.

Screen reader:

- Annual Report dot t x t - Notepad
- Annual Report

Confirm the wording with a live run: buildTutorials -live.

### Step 10: Down Arrow

Down Arrow reads on, line by line, with nothing between me and the words.

Screen reader:

- Summary

Confirm the wording with a live run: buildTutorials -live.

### Step 11: Control+A

Control plus A selects all of it and Control plus C copies it, ready to paste into an email.

Screen reader:

- selected

Confirm the wording with a live run: buildTutorials -live.

### Step 12

The same file can go to a text-to-speech program, which reads it aloud as an audio book of the document.

### Step 13

Or to a braille display, where clean text without formatting marks reads more easily.

### Step 14: Alt+U

Use configuration, U, remembers these boxes, so the next time the dialog opens ready for plain text.

Screen reader:

- Use configuration check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 15

To go back to web pages, clear Plain text; Default settings, D, clears every box at once.

### Step 16

The text file is written in UTF-8, so accented letters and symbols arrive intact in any modern program.

### Step 17

Each paragraph is one line, so a text reader that pauses at line ends pauses at paragraph ends, not mid-sentence.

### Step 18

A heading stands on a line of its own, with a blank line around it, so a listener hears where a section begins.

### Step 19

List items keep their marks, a dash for a bullet and a number for a numbered list, so their order is still heard.

### Step 20

Plain text of a whole folder works the same way: one text file for each document, ready to join into one long reading.

### Step 21

Plain text is the form to choose whenever the words matter and the layout does not.

### Step 22

One document, two forms: a web page to navigate, or plain text to hear or paste. The next walk converts slides and spreadsheets.

**Something to try:** Make plain text of one document, and paste it into an email.

## 6 - Convert Slides and Spreadsheets

Converting PowerPoint slides and Excel workbooks into web pages, and how each is laid out for reading.

**Before you start:** A presentation or a workbook of your own, and Microsoft Office installed.

### Step 1

One want: the slides from a talk I missed, and the budget workbook shared with it, both readable without opening PowerPoint or Excel.

### Step 2: Alt+Control+2

I open the dialog. Source files can hold several files at once, separated by spaces.

Screen reader:

- 2htm dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3

I type both names, each in quotation marks because they contain spaces.

Screen reader:

- quote Talk dot pptx quote quote Budget dot xlsx quote

Confirm the wording with a live run: buildTutorials -live.

### Step 4: Enter

Enter converts both. PowerPoint and Excel read them in the background.

Screen reader:

- 2 files converted

Confirm the wording with a live run: buildTutorials -live.

### Step 5

I open the slides' page. Each slide is a section with its title as a heading.

Screen reader:

- Talk, heading level 1

Confirm the wording with a live run: buildTutorials -live.

### Step 6: H

H moves slide by slide.

Screen reader:

- Slide 2, Our goals, heading level 2

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Down Arrow

Under each heading come the slide's words, then its speaker notes, so what the presenter said is there too.

Screen reader:

- Notes: start with the year in review

Confirm the wording with a live run: buildTutorials -live.

### Step 8: G

Pictures on a slide read their descriptions, if the author wrote them.

Screen reader:

- graphic, Map of the new offices

Confirm the wording with a live run: buildTutorials -live.

### Step 9

Now the workbook's page. Each sheet becomes a section with the sheet's name as its heading.

Screen reader:

- Budget, heading level 1

Confirm the wording with a live run: buildTutorials -live.

### Step 10: T

Under it, the sheet's cells are a table, so the reader's table keys move by row and column, and each cell is read with its column's heading.

Screen reader:

- table with 5 columns and 12 rows

Confirm the wording with a live run: buildTutorials -live.

### Step 11: Alt+Control+Down Arrow

Alt plus Control plus Down Arrow moves down a column.

Screen reader:

- Amount, 1,200

Confirm the wording with a live run: buildTutorials -live.

### Step 12

Formulas arrive as the values they show, which is what a reader of the page wants.

### Step 13

A very wide sheet reads best with the screen reader's table commands, rather than line by line.

### Step 14

A chart in a slide reads the description its author wrote; with none, it reads as a graphic with no name, which is worth telling the author.

### Step 15

Slides with several text boxes keep them in reading order, the order a screen reader meets them in PowerPoint itself.

### Step 16

A workbook's hidden sheets are not converted, so the page holds only what its author shows.

### Step 17

Merged cells keep their place in the table, so a heading spread across several columns still sits above them.

### Step 18

Very long sheets make long tables; the reader's table commands, such as moving to the first or last row, save a long walk.

### Step 19

A deck converted the night before a talk can be studied on a phone, where opening PowerPoint is harder.

### Step 20

A presentation's title slide becomes the page's first heading, so the page opens by saying what talk it is.

### Step 21: Shift+H

And a workbook's sheets keep their order, so Shift plus H, back by heading, retraces the sheets in the order the author made them.

Screen reader:

- Budget, heading level 1

Confirm the wording with a live run: buildTutorials -live.

### Step 22

Slides by heading, sheets as tables: two kinds of Office file, read without Office open. The next walk converts a whole folder.

**Something to try:** Convert a slide deck, then move through it by heading, one slide at a time.

## 7 - Convert a Whole Folder

Converting every document in a folder at once, keeping the pages together, and keeping a log of the run.

**Before you start:** A folder holding several documents, such as meeting papers.

### Step 1

One want: every document from a week of meetings, converted in one run, the pages kept together in their own folder.

### Step 2: Alt+Control+2

I open the dialog. In Source files, a folder's path is enough: 2htm converts every document in it that it knows how to read.

Screen reader:

- 2htm dialog
- Source files edit

Confirm the wording with a live run: buildTutorials -live.

### Step 3

I type the folder's path.

Screen reader:

- C colon backslash Meetings backslash October

Confirm the wording with a live run: buildTutorials -live.

### Step 4

A pattern narrows it: star dot docx would take only the Word documents.

### Step 5: Alt+O

I want the pages in a folder of their own, so I go to Output directory.

Screen reader:

- Output directory edit

Confirm the wording with a live run: buildTutorials -live.

### Step 6

I type a new folder, Pages, beside the meeting folder.

Screen reader:

- C colon backslash Meetings backslash Pages

Confirm the wording with a live run: buildTutorials -live.

### Step 7: Alt+L

Log session, L, keeps a log of the whole run beside the pages.

Screen reader:

- Log session check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 8: Alt+V

View output, V, opens the Pages folder at the end.

Screen reader:

- View output check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 9: Enter

Enter. The Pages folder does not exist yet, so 2htm asks before making it; Yes is the default.

Screen reader:

- Create the output directory?
- Yes button

Confirm the wording with a live run: buildTutorials -live.

### Step 10: Enter

Enter, and the run begins. A small window counts the files, so a long run is never silent.

Screen reader:

- Converting 3 of 9: Minutes dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 11

The result box gives the totals.

Screen reader:

- 9 files converted

Confirm the wording with a live run: buildTutorials -live.

### Step 12

If two documents share a name, such as Agenda dot docx and Agenda dot pdf, the second page takes its kind into its name, so neither replaces the other.

### Step 13

Running again tomorrow skips the documents whose pages exist, so only new ones are converted; Force replacements converts them all again.

### Step 14

If one document cannot be read, such as a file locked by a password, the others still convert, and the result names the one that failed.

Screen reader:

- 8 files converted, 1 failed

Confirm the wording with a live run: buildTutorials -live.

### Step 15: Alt+U

Use configuration remembers these folders and boxes, so tomorrow's run is one key.

Screen reader:

- Use configuration check box, checked

Confirm the wording with a live run: buildTutorials -live.

### Step 16

Subfolders are not searched unless named, so a run converts only what you meant; name each subfolder, or use a pattern, to include them.

### Step 17

The log is a plain text file listing every document, what became of it, and any error in full, which is what to send if something seems wrong.

### Step 18

The progress window can be left alone; the run carries on, and its count says how far it has come.

Screen reader:

- Converting 7 of 9: Budget dot xlsx

Confirm the wording with a live run: buildTutorials -live.

### Step 19

Opening the Pages folder lists the converted pages with the same names as their documents, so finding one is a matter of its first letter.

### Step 20

Office files are converted one at a time, in Office's own hidden copy, so Word or Excel already open for your own work is left alone.

### Step 21

A long run can be stopped from the progress window; documents finished so far keep their pages.

### Step 22

A week's papers, converted together, one folder of pages. The last task walk does it from File Explorer and the command line.

**Something to try:** Convert a folder twice, the second time without Force replacements, and hear the documents already converted skipped.

## 8 - Convert from Explorer and the Command Line

Converting without the dialog: from File Explorer's context menu, and from a command prompt or batch file, with the same options.

**Before you start:** A document in File Explorer, and a command prompt.

### Step 1

One want: to convert a document the moment I find it, and to convert a folder every week without opening anything.

### Step 2

In File Explorer, I am on a document.

Screen reader:

- Notes dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 3: Shift+F10

Shift plus F10 opens its context menu.

Screen reader:

- Context menu

Confirm the wording with a live run: buildTutorials -live.

### Step 4: 2

2 is Convert via 2htm, and the page appears beside the document.

Screen reader:

- 2htm: 1 file converted

Confirm the wording with a live run: buildTutorials -live.

### Step 5

At a command prompt, 2htm followed by a file's name converts it.

Screen reader:

- 2htm Notes dot docx

Confirm the wording with a live run: buildTutorials -live.

### Step 6

Every box in the dialog has an option. Dash p makes plain text; dash s strips images; dash f replaces existing pages; dash o sends output to a folder.

### Step 7

So one line converts a whole folder into a Pages folder, replacing last week's pages.

Screen reader:

- 2htm dash f dash o Pages C colon backslash Meetings backslash October

Confirm the wording with a live run: buildTutorials -live.

### Step 8

Quotation marks keep a path with spaces in one piece.

### Step 9

Saved as a batch file, that line runs in a moment, or on a schedule with Windows Task Scheduler.

### Step 10

Dash u reads the settings saved by the dialog's Use configuration box, so a workflow tried in the dialog runs the same from a batch file.

### Step 11

Dash l writes the session log beside the output; a log is always kept in your own application data as well.

### Step 12

When it finishes, 2htm returns an exit code a batch file can test: zero when every file converted, one when a file could not be, or none were found.

Screen reader:

- if errorlevel 1 echo Some files were not converted

Confirm the wording with a live run: buildTutorials -live.

### Step 13

Dash h lists every option, and dash v says the version.

Screen reader:

- 2htm dash h

Confirm the wording with a live run: buildTutorials -live.

### Step 14

Dash g opens the dialog, for anyone who starts at the prompt but prefers the fields.

### Step 15

Patterns work at the prompt as in the dialog: star dot pdf converts every PDF in the current folder.

Screen reader:

- 2htm star dot pdf

Confirm the wording with a live run: buildTutorials -live.

### Step 16

Several patterns can be given at once, separated by spaces, to convert two kinds of file in one run.

### Step 17

A Task Scheduler entry that runs the batch file each Monday morning keeps a folder of pages current with no effort.

### Step 18

Dash dash view output opens the output folder at the end, as the dialog's box does.

### Step 19

Running 2htm with no arguments at all opens the dialog, so a shortcut to the program works without options.

### Step 20

Everything learned in the dialog carries over to the prompt, and back again.

### Step 21

From the menu, from a prompt, or on a schedule: the same conversions wherever the work is. The last walk gathers it all together.

**Something to try:** Write a one-line batch file that converts a folder into a Pages folder, and run it.

## 9 - Conclusion

What the walks covered, a glossary in two voices, and every way to get help with 2htm.

**Before you start:** Nothing to set up; this walk only listens.

### Step 1

That is 2htm. Here is what each walk gave you.

### Step 2

Walk one, the dialog and a letter for everything. Walk two, installing and the three ways to start. Walks three to six: a Word report, a PDF, plain text, and slides and spreadsheets.

### Step 3

Walk seven, a whole folder at once. Walk eight, File Explorer and the command line.

### Step 4

Now a short glossary. I say the term; the other voice says what it means.

### Step 5

Web page.

Screen reader:

- A document a browser shows, with headings, lists and tables a screen reader can move through by key.

### Step 6

Plain text.

Screen reader:

- Words alone, with no formatting, for listening, braille or pasting into an email.

### Step 7

Heading.

Screen reader:

- A title in the document's outline; the H key moves from one to the next.

### Step 8

Scanned PDF.

Screen reader:

- A PDF made of pictures of pages, with no text inside until text recognition reads it.

### Step 9

Output directory.

Screen reader:

- The folder where converted files are written; blank means beside each document.

### Step 10: F1

Help is never more than a key away. In the dialog, F1 shows Help, and Shift plus F1 says the tip for the control you are on.

Screen reader:

- 2htm Help dialog

Confirm the wording with a live run: buildTutorials -live.

### Step 11: Alt+G

F7 lists the dialog's controls. Guide, Alt plus G, opens the full guide and returns to the dialog as you left it.

Screen reader:

- Guide button

Confirm the wording with a live run: buildTutorials -live.

### Step 12

F11 checks for a newer version, and the 2htm page on GitHub has the latest release and a place to report a problem.

### Step 13

Thank you for listening. Happy converting.

**Something to try:** Convert the document you read most often, and keep its web page.

<!-- walkthrough ends -->

## 1. Where to Go Next

Press F1 for the guide, Alt+Shift+H for the hotkey list, and Alt+F10 for every
command in one window.
