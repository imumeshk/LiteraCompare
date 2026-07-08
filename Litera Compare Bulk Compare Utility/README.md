# Litera Compare Bulk Compare Utility

The Litera Compare Bulk Compare Utility is a powerful script designed to automate the process of comparing multiple documents using the Litera Compare engines. It features a graphical user interface (GUI) built in PowerShell that simplifies complex batch comparison tasks.

## Functionality

This utility leverages the Litera Compare command-line executables (e.g., `lcp_auto.exe`, `lcp_main.exe`) to run automated, unattended comparisons. It allows users to select an original document (or a folder of original documents) and compare it against modified documents, generating redline output files (such as PDFs or Word documents) automatically.

This significantly reduces the manual effort required to perform large-scale document comparisons, making it ideal for legal, financial, and administrative professionals dealing with numerous revisions.

## Comparison Modes

The tool provides three primary functions/modes of operation:

1. **Single (One Original vs Folder of Modified)**
   Compares a single selected original document against every document inside a selected modified folder. This is useful when you have one base document and several different modified versions of it that you want to evaluate.

2. **Bulk (Folder vs Folder - Sequential Match)**
   Compares the contents of an Original folder against the contents of a Modified folder sequentially (e.g., the 1st file in the Original folder is compared against the 1st file in the Modified folder, the 2nd vs 2nd, etc.).

3. **Exact (Folder vs Folder - Exact Match)**
   Compares files in the Original folder against files in the Modified folder that share the exact same file name.

## Key Features & Options

- **Comparison Engine Selection**: Choose between `Auto`, `Word`, `PowerPoint`, `Excel`, or `PDF` engines depending on the document types being compared.
- **Customizable Output Format**: Set the desired output format for the redline document (e.g., `.pdf`, `.docx`).
- **File Prefixing**: Automatically prepend a prefix (e.g., `redline_`) to the generated output files to easily identify them in the destination folder.
- **Track Changes**: An option to output the comparison as a document with Track Changes enabled instead of a static redline.
- **Advanced Execution Options**: Includes toggles for showing the visible UI during comparison, prompting on errors, and rendering only pages with redline changes (`RedlinePagesOnly`).
