# ASSIGNMENT TASKS
 <!-- File Manipulation Task
 1. Write a script to extract all text and images from the provided PDF/DOCX. Ensure that all
 images are saved to the disk.
 * Pay attention to how the files/folders are named during this process.
 2. Foreachparagraph in the PDF/DOCX, extract the following details:
 ● Textcontent
 ● Fonttype
 ● Fontsize
 ● Styling elements (such as italics, bold, etc.)
 ● Textcolor
 3. Convert the text of each extracted paragraph to UPPERCASE. Subsequently, compile all the
 UPPERCASE paragraphs into a new PDF/DOCX, maintaining the original formatting (font type
 and styling) as closely as possible.
 4. Write ascript to extract all text and images from the provided PPTX and then translate all the
 text in file to English and then append the translated text under the original text back in
 slides, please try to keep the font size as reasonable as possible.
 Note:--
 If possible, use container for deployment
 Submit code via Git (Gitlab or Github is highly recommended)
 Reference
 [1] A high performance Python library for data extraction, analysis, conversion & manipulation of PDF
 (and other) documents. https://pypi.org/project/PyMuPDF/
 [2] A flexible free and unlimited python tool to translate between different languages in a simple way
 using multiple translators. https://pypi.org/project/deep-translator -->
## File Manipulation Task
## Todo List

- [x] Write a script to extract all text and images from the provided PDF/DOCX.
    - [x] PDF (PyMuPDF) - extracted to folder with metadata
        - [x] Text content
        - [x] Font name
        - [x] Font size
        - [x] Text color
        - [x] Images
    - [x] DOCX (python-docx) - extracted to metadata
        - [x] Text content
        - [x] Font type
        - [x] Font size
        - [x] Text color 
        - [x] Images
        - Note: Support for font using from "Style" or "Default Style" is not supported by the current implementation.
       
- [x] For each paragraph in the PDF/DOCX, extract the following details.
    - [x] Text content
    - [x] Font type
        - [x] PDF
        - [x] DOCX (Only for manually specified fonts)
        - Note: Support for font using from "Style" or "Default Style" is not supported by the current implementation.
    - [x] Font size
    - [x] Styling elements (such as italics, bold, etc.)
        - [x] Italic
        - [x] Bold
        - [ ] Underline
        - [ ] Strikethrough
    - [x] Text color
- [x] Convert the text of each extracted paragraph to UPPERCASE.
    - [x] Compile all the UPPERCASE paragraphs into a new PDF/DOCX.
    - [x] Maintain the original formatting (font type and styling) as closely as possible.
        - [x] PDF
            - Note: The current implementation does not support font type.
        - [x] DOCX
            - [ ] Tables
- [ ] Translate all text in the file to English and append the translated text under the original text in slides.
- [x] Save all images to the disk, paying attention to how the files/folders are named during this process.
- [ ] Deploy the application using a container.
- [ ] Submit code via Git (Gitlab or Github is highly recommended).
