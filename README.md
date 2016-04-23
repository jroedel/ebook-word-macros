# ebook-word-macros
Word macros to help format ebooks. Tested in Word 2016

#Installation
**Import macro file:**
1. In the Word ribbon, goto View->Macros->Edit (or Create if there are no pre-existing macros)
2. In the Project Explorer, right click "Normal"->Import File...
3. Find the .bas file

**Import ribbon customizations for easy access to main macros:**
1. Right click a blank space on the Word ribbon -> Customize ribbon
2. Click Import/Export->Import customization file
3. Find the .exportedUi file

#Macros
* **Start Book from Wpd** - I use for fixing simple formatting problems coming from a WordPerfect document
* **Fix Blank Paragraphs** - Used for substituting blank paragraphs (to separate paragraphs) with "Space after paragraph". Ebooks prefer this kind of spacing so that pages never start with blank space on top
* **Put Manual Page Break Before Headings** - Places a manual page break before Headings 1 and 2
* **Add Chapter Hyperlinks** - Creates bookmarks for all Heading 2's and creates an index immediately after each Heading 1
* **Place Index At End** - Creates a Word-generated Index (according to headings) at the end of the document
* **Convert Selection to Verse Styles** - Normally my ebook paragraphs have space added after. This creates and applies a style that eliminates the space between paragraphs for poetry/lyrical text.
* **Convert Lists to Text** - Bullet points and numbering doesn't come out well when parsed into ebook formats. This macro converts Bullets and Numbering to plain text.
* **Reapply Heading Styles** - Sometimes Word gets confused. This searches for all the paragraphs that have Outline Level 1 or 2 and applies the appropriate Heading style.
* **Eliminate Multiple Spaces** - Replaces any occurance of a double space with a single space. Spaces should never be used for indents.
* **Replace Tabs with Space** - Tabs are evil.
* **Find Next Separated Paragraph** - When converting PDFs to Word, often the paragraphs get cut off in the middle. This macro applies a wildcard search to search out broken paragraphs. No changes are made, just finds the next instance.
* **What Chr is this?** - Gives the ASCII code of the first selected character
* **Bold All Caps Text** - I think ALL CAPS is ugly. This changes the case of ALL CAPS text and bolds it instead
* **Change All Caps to Sentence Case** - Careful with Roman Numerals, I tried to exclude them but sometimes they get lowered when they're part of a longer heading
* **Find Next Hard Hyphenated Word** - OCR and PDF conversions leave some words with hyphens in between. This does a Find to try to hunt them down
* **Fix Em Dashes** - In Spanish, clarification clauses -parts of sentences that explain the preceding idea or concept- should be separated by long dashes called Em Dashes. This macro does a find and replace to fix these.
* **Fix Quotes** - Older versions of Word or converted files sometimes confuse if a quote is opening or closing. This macro mostly fixes this by searching " or ' and replacing them with " or ' respectively. It just works.

##Word Replacement Files
The files I work with have a very specific vocabulary. I have two files where I store common word replacements (for example: sicologia to psicología). One file is a straight case-sensitive match-whole-word find and replace. The other file is a wildcard based search. The macros look for these two files in the same directory as the file that contains the macro (probably in the directory of the Normal template). The name of the file can be specified in the const variable at the top of the .bas file. For this functionality I have 4 macros:
* Open Word Replacement File - Opens the file in notepad++ or notepad
* Open Wildcard Replacement File - I lied. I haven't written this macro yet. #todo
* Add Selection to Word Replacement - takes the selected characters and prompts the user for the replacement of those characters. Adds this as a record in the Replacemente File.
* Common Word Replacements - does three things:
  1. Offers to de-capitalize the Spanish months of the year (as opposed to English months which are always capitalized)
  2. Executes the find/replace of the Word Replacement File
  3. Executes the find/replace of the Wildcard Replacement File
