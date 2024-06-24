import os
import json
from docx import Document
from docx.shared import Pt, RGBColor
from io import BytesIO
import fitz  # PyMuPDF

class FileManipulator:
    def __init__(self):
        pass

    def extract_text_images_from_pdf(self, pdf_path, output_folder):
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        doc = fitz.open(pdf_path)
        metadata = {"pages": []}
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # Extract text with styling information
            page_blocks = page.get_text("dict")["blocks"]
            text_metadata = []
            
            for block in page_blocks:
                if block["type"] == 0:  # Text block
                    for line in block["lines"]:
                        for span in line["spans"]:
                            bbox = span["bbox"]
                            text_meta = {
                                "text": span["text"],
                                "font_size": span["size"],
                                "font_color": span["color"],
                                "font_name": span["font"],
                                "bbox": bbox,
                                "bold": span["flags"] & 1 
                                  != 0 or ("bold" in span["font"].lower()) ,  
                                "italic": span["flags"] & 2 != 0 or ("italic" in span["font"].lower()), 
                            }
                            text_metadata.append(text_meta)
            
            # Extract images (similar to your current implementation)
            images = page.get_images(full=True)
            image_metadata = []
            for img_index, img in enumerate(images):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_path = os.path.join(output_folder, f"page_{page_num + 1}_image_{img_index + 1}.{image_ext}")
                with open(image_path, "wb") as img_file:
                    img_file.write(image_bytes)
                image_metadata.append({
                    "index": img_index,
                    "path": image_path,
                    "ext": image_ext
                })
            
            metadata["pages"].append({
                "page_num": page_num + 1,
                "text_instances": text_metadata,
                "images": image_metadata
            })
        
        with open(os.path.join(output_folder, "metadata.json"), "w") as meta_file:
            json.dump(metadata, meta_file, indent=4)
 
    def extract_text_images_from_docx(self, docx_path, output_folder):
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        doc = Document(docx_path)
        metadata = {"paragraphs": [], "images": []}
        
        for para_num, para in enumerate(doc.paragraphs):
            para_data = {
                "para_num": para_num + 1,
                "text": para.text,
                "runs": []
            }
            for run in para.runs:
                run_data = {
                    "text": run.text,
                    "bold": run.bold,
                    "italic": run.italic,
                    "font_name": run.font.name,
                    "font_size": run.font.size.pt if run.font.size else None,
                    "color": run.font.color.rgb if run.font.color else None
                }
                para_data["runs"].append(run_data)

            metadata["paragraphs"].append(para_data)
            with open(os.path.join(output_folder, f"paragraph_{para_num + 1}.txt"), "w") as text_file:
                text_file.write(para.text)
        
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                img = rel.target_part.blob
                img_ext = rel.target_part.content_type.split('/')[-1]
                img_path = os.path.join(output_folder, f"image_{rel.target_ref.split('/')[-1]}.{img_ext}")
                with open(img_path, "wb") as img_file:
                    img_file.write(img)
                metadata["images"].append({
                    "path": img_path,
                    "ext": img_ext
                })
        
        with open(os.path.join(output_folder, "metadata.json"), "w") as meta_file:
            json.dump(metadata, meta_file, indent=4)

    def convert_text_to_uppercase(self, pdf_path, output_path, type):
        if type == "pdf":
            self.recreate_pdf("output_pdf", output_path, lambda text: text.upper())
        elif type == "docx":
            self.recreate_docx("output_docx", output_path, lambda text: text.upper())
    
    def recreate_docx(self, output_folder, output_path, content_processor = None):
        with open(os.path.join(output_folder, "metadata.json"), "r") as meta_file:
            metadata = json.load(meta_file)
        
        doc = Document()
        
        for para_meta in metadata["paragraphs"]:
            para = doc.add_paragraph()
            for run_meta in para_meta["runs"]:
                if content_processor:
                    run = para.add_run(content_processor(run_meta["text"]))
                else:
                    run = para.add_run(run_meta["text"])
                run.bold = run_meta["bold"] if run_meta["bold"] else False
                run.italic = run_meta["italic"] if run_meta["italic"] else False
                if run_meta["font_size"]:
                    run.font.size = Pt(run_meta["font_size"])
                if run_meta["color"]:
                    color_int = run_meta["color"]
                    r = color_int[0]
                    g = color_int[1]
                    b = color_int[2]
                    run.font.color.rgb = RGBColor(r, g, b)
                run.font.name = run_meta["font_name"]
        
        for img_meta in metadata["images"]:
            with open(img_meta["path"], "rb") as img_file:
                doc.add_picture(img_file, width=Pt(300))
        
        doc.save(output_path)

    def recreate_pdf(self, output_folder, output_path, content_processor = None):
        with open(os.path.join(output_folder, "metadata.json"), "r") as meta_file:
            metadata = json.load(meta_file)
        
        doc = fitz.open()
        
        for page_meta in metadata["pages"]:
            page = doc.new_page()
            
            for text_meta in page_meta["text_instances"]:
                bbox = fitz.Rect(text_meta["bbox"])
                r = (text_meta["font_color"] >> 16) & 0xff
                g = (text_meta["font_color"] >> 8) & 0xff
                b = text_meta["font_color"] & 0xff
                color = [r/255, g/255, b/255]
                page.insert_text((bbox.x0, bbox.y0),
                                 content_processor(text_meta["text"]) if content_processor else text_meta["text"]
                                 , fontsize=text_meta["font_size"], color=color)
            for img_meta in page_meta["images"]:
                image_rect = fitz.Rect(72, 72 + img_meta["index"] * 200, 300, 400 + img_meta["index"] * 200)
                page.insert_image(image_rect, filename=img_meta["path"])
        
        doc.save(output_path)


if __name__ == "__main__":
    manipulator = FileManipulator()

    # Extract text and images from PDF
    manipulator.extract_text_images_from_pdf("docx_mock_file.pdf", "output_pdf")

    # Extract text and images from DOCX
    manipulator.extract_text_images_from_docx("docx_mock_file.docx", "output_docx")

    # Convert text to uppercase and compile into new DOCX
    manipulator.convert_text_to_uppercase("docx_mock_file.docx", "uppercase_docx.docx",type="docx")

    # Convert text to uppercase and compile into new PDF
    manipulator.convert_text_to_uppercase("docx_mock_file.pdf", "uppercase_pdf.pdf", type="pdf")

    # Recreate DOCX from extracted data
    manipulator.recreate_docx("output_docx", "recreated_docx.docx")

    # Recreate PDF from extracted data
    manipulator.recreate_pdf("output_pdf", "recreated_pdf.pdf")
