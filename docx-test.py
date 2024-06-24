from docx import Document
import os


def get_effective_font_name(run, doc):
    # Check if the run has a font name set
    if run.font.name:
        return run.font.name
    # Check the paragraph style if run.font.name is None
    elif run._parent.style.font.name:
        return run._parent.style.font.name
    # Fallback to the document's 'Normal' style
    elif doc.styles['Normal'].font.name:
        return doc.styles['Normal'].font.name
    # Default font name if none is set
    else:
        return "Default Font"


def read_docx_with_details_and_extract_images(file_path):
    doc = Document(file_path)
    detailed_text = []

    # Ensure the /img directory exists
    img_dir = "./img"
    os.makedirs(img_dir, exist_ok=True)
    
    # Extract text with styles
    for para in doc.paragraphs:
        for run in para.runs:
            run_text = apply_formatting(run, run.text)
            font_name = get_effective_font_name(run,doc)
            font_size = run.font.size.pt if run.font.size else "Default Size"
            text_color = run.font.color.rgb if run.font.color and run.font.color.rgb else "Default Color"
            details = f"Text: {run_text}, Font: {font_name}, Size: {font_size}, Color: {text_color}"
            detailed_text.append(details)
        detailed_text.append('\n')
    
    # Extract and save images
    image_count = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image = rel.target_part.blob
            image_path = os.path.join(img_dir, f"image{image_count}.png")
            with open(image_path, "wb") as img_file:
                img_file.write(image)
            image_count += 1
    
    return '\n'.join(detailed_text)

def apply_formatting(run, run_text):
    formatting_actions = {
        "bold": lambda text: f"**{text}**",
        "italic": lambda text: f"*{text}*",
        "underline": lambda text: f"__{text}__",
    }

    for style, action in formatting_actions.items():
        if getattr(run, style):
            run_text = action(run_text)

    return run_text

# Example usage
print(read_docx_with_details_and_extract_images("plain.docx"))