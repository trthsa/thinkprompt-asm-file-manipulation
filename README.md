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
# Introduction
In this assignment, I will be working on various file manipulation tasks using Python. The main goal is to extract text and images from PDF/DOCX files, manipulate the extracted data, and save it in a new format while maintaining the original formatting as closely as possible. Additionally, I will also be translating the text in the file to English and appending it under the original text in slides. Finally, I will containerized the application.

## Zero Setup Deployment

Click the button below to start a new ready-to-go development environment:

[![Open in Gitpod](https://gitpod.io/button/open-in-gitpod.svg)](https://gitpod.io/#https://github.com/trthsa/thinkprompt-asm-file-manipulation)

## Requirements
- For local runs Python 3.12+ or Docker v24+ is required for running the application in a container.
- See the requirements.txt file for the required packages.

## Installation
1. Clone the repository
2. Install the required packages using pip:
```bash
pip install -r requirements.txt
```
3. Run the server.py file to start the server on 0.0.0.0:5000 or localhost:5000 or 127.0.0.1:5000

```bash
python server.py
```
4. The output files will be saved in the `results` folder.
## Running the application using Docker
1. Clone the repository
2. Build the Docker image using the following command:
```bash
docker build -t file-manipulator .
```
3. Run the Docker container using the following command:
```bash
docker run -p 5000:5000 file-manipulator
```
4. The output files will be saved in the `results` folder.


## Guide
- The server.py file contains the main code for the application.
- The `results` folder contains the output files.
- The `uploads` folder contains the uploaded files.
-  [<img src="https://run.pstmn.io/button.svg" alt="Run In Postman" style="width: 128px; height: 32px;">](https://app.getpostman.com/run-collection/12812349-6fd3ee6d-7481-406c-b8b5-3f06d33c0bfe?action=collection%2Ffork&source=rip_markdown&collection-url=entityId%3D12812349-6fd3ee6d-7481-406c-b8b5-3f06d33c0bfe%26entityType%3Dcollection%26workspaceId%3D08e923f8-b871-4725-a244-0f21f4d14775#?env%5BFM-API-ENV%5D=W3sia2V5IjoiaG9zdCIsInZhbHVlIjoiaHR0cDovLzEyNy4wLjAuMTo1MDAwIiwiZW5hYmxlZCI6dHJ1ZSwidHlwZSI6ImRlZmF1bHQifV0=) for more details on the API endpoints and pre-crafted examples. or import `FileManipulator API.postman_collection.json` to your local Postman.
## File Manipulation Task
## Todo List
> [!NOTE]
> The following tasks are based on the requirements provided in the assignment. The tasks are divided into sub-tasks to make it easier to track the progress of the project. The tasks are also marked as completed if the implementation is not possible due to limitations (with a note explaining the limitation)

> [!IMPORTANT]
> The project doesn't support any security features like authentication, authorization, or encryption. The project is designed to be a simple file manipulation tool. The project is not designed to be used in a production environment. Please use the project with caution.

- [x] Write a script to extract all text and images from the provided PDF/DOCX.
    - [x] PDF (PyMuPDF) - extracted to folder with metadata (JSON format)
        - [x] Text content
        - [x] Font name
        - [x] Font size
        - [x] Text color
        - [x] Images
    - [x] DOCX (python-docx) - extracted to metadata (JSON format)
        - [x] Text content
        - [x] Font type
        - [x] Font size
        - [x] Text color 
        - [x] Images
        - Note: DOCX Support for font using from "Style" or "Default Style" is not supported by the current implementation.
       
- [x] For each paragraph in the PDF/DOCX, extract the following details.
    - [x] Text content
    - [x] Font type
        - [x] PDF
        - [x] DOCX (Only for manually specified fonts)
        - Note 1: Support for font using from "Style" or "Default Style" is not supported by the current implementation.
        - Note 2: PDF Support is not supported by the current implementation.
    - [x] Font size
    - [x] Styling elements (such as italics, bold, etc.)
        - [x] Italic
        - [x] Bold
        - Note: PDF Support styling is not supported by the current implementation.
    - [x] Text color
- [x] Convert the text of each extracted paragraph to UPPERCASE.
    - [x] Compile all the UPPERCASE paragraphs into a new PDF/DOCX.
    - [x] Maintain the original formatting (font type and styling) as closely as possible.
        - [x] PDF
            - Note: The current implementation does not support font type and styling.
        - [x] DOCX
- [x] Translate all text in the file to English and append the translated text under the original text in slides.
- [x] Save all images to the disk, paying attention to how the files/folders are named during this process.
- [x] Deploy the application using a container.
- [x] Submit code via Git (Gitlab or Github is highly recommended).
