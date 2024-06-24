import pymupdf
import os

from tree_docx import build_pdf_tree, display_tree, recreate_docx_from_tree

def read_pdf_with_details_and_extract_images(file_path):
	doc = pymupdf.open(file_path) 
	pdfByptes = doc.convert_to_pdf()
	pdf= pymupdf.open("pdf",pdfByptes)
	img_dir = "./img"
	os.makedirs(img_dir, exist_ok=True)  # Ensure the img directory exists
	
	for page_num, page in enumerate(pdf):  # iterate the document pages
		blocks = page.get_text("dict")["blocks"]
		for block in blocks:
			if "lines" in block:  # Ensure block contains lines
				for line in block["lines"]:
					for span in line["spans"]:  # Iterate through each span in the line
						text = span["text"]
						font_name = span["font"]
						font_size = span["size"]
						color_int = span["color"]
						r = (color_int >> 16) & 0xff
						g = (color_int >> 8) & 0xff
						b = color_int & 0xff
						color_hex = f"#{r:02x}{g:02x}{b:02x}"
						print(f"Text: {text}, Font: {font_name}, Size: {font_size}, Color: {color_hex}")
		
		# Extract and save images
		image_list = page.get_images(full=True)
		for image_index, img in enumerate(image_list):
			xref = img[0]
			base_image = pdf.extract_image(xref)
			image_bytes = base_image["image"]
			image_path = os.path.join(img_dir, f"page_{page_num}_image_{image_index}.png")
			with open(image_path, "wb") as img_file:
				img_file.write(image_bytes)
# Usage
pdf_tree = build_pdf_tree("plain.pdf")
display_tree(pdf_tree)
# Assuming `doc_tree` is the tree structure created from the original document
new_doc = recreate_docx_from_tree(pdf_tree)
new_doc.save("recreated_docx_file_pdf.docx")
# read_pdf_with_details_and_extract_images("docx_mock_file_test.pdf")