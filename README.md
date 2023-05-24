# BlueBeam-eX-Excel-Submittal-from-Folder-Structure

This code is written in VBA for EXCEL.
It will use the Bluebeam Script Engine (ScriptEngine.exe) through a Scripting.Shell interface.

What it does:
I takes a folder that has multiple subfolders that each contain PDFs (and Word Docs) and merges them into a single PDF, perserving the folder structure in the Bookmarks and page order.

I built this so I could put together Engineering Drawing submittals that have product data, Schematics, and other supporting documentation.

The code is split into 3 main Subs:
1- Section Generator: This will make a Word docx file for each folder (and subfolder) which will act as a Section page.
2- Word 2 PDF: This will take all the Word documents in the folder structure and convert them to PDF with the same name.
3- Submittal Builder: This will merge all of the PDFs together in such a way that results in a Single PDF with Page order and Bookmarks that mirror the folder structure higherarchy.
