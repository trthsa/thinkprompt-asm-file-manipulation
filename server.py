from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename
from file_manipulator import FileManipulator
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'results'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

manipulator = FileManipulator()

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    
    return jsonify({"file_path": file_path}), 200

@app.route('/extract_text_images_from_pdf', methods=['POST'])
def extract_text_images_from_pdf():
    pdf_path = request.json['pdf_path']
    output_folder = os.path.join(app.config['OUTPUT_FOLDER'], 'pdf_output')
    manipulator.extract_text_images_from_pdf(pdf_path, output_folder)
    metadata_path = os.path.join(output_folder, 'metadata.json')
    return jsonify({"status": "success", "metadata_path": metadata_path}), 200

@app.route('/extract_text_images_from_docx', methods=['POST'])
def extract_text_images_from_docx():
    docx_path = request.json['docx_path']
    output_folder = os.path.join(app.config['OUTPUT_FOLDER'], 'docx_output')
    manipulator.extract_text_images_from_docx(docx_path, output_folder)
    metadata_path = os.path.join(output_folder, 'metadata.json')
    return jsonify({"status": "success", "metadata_path": metadata_path}), 200

@app.route('/convert_text_to_uppercase', methods=['POST'])
def convert_text_to_uppercase():
    file_path = request.json['file_path']
    file_type = request.json['file_type']
    output_file = os.path.join(app.config['OUTPUT_FOLDER'], f'uppercase_output.{file_type}')
    manipulator.convert_text_to_uppercase(file_path, output_file, file_type)
    return jsonify({"status": "success", "output_path": output_file}), 200

@app.route('/recreate_docx', methods=['POST'])
def recreate_docx():
    output_folder = request.json['output_folder']
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'recreated_docx.docx')
    manipulator.recreate_docx(output_folder, output_path)
    return jsonify({"status": "success", "output_path": output_path}), 200

@app.route('/recreate_pdf', methods=['POST'])
def recreate_pdf():
    output_folder = request.json['output_folder']
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'recreated_pdf.pdf')
    manipulator.recreate_pdf(output_folder, output_path)
    return jsonify({"status": "success", "output_path": output_path}), 200

@app.route('/translate_text_in_pptx', methods=['POST'])
def translate_text_in_pptx():
    input_pptx_path = request.json['input_pptx_path']
    output_pptx_path = os.path.join(app.config['OUTPUT_FOLDER'], 'translated_pptx.pptx')
    from_lang = request.json.get('from_lang', 'en')
    to_lang = request.json.get('to_lang', 'vi')
    manipulator.translate_text_in_pptx(input_pptx_path, output_pptx_path, from_lang, to_lang)
    return jsonify({"status": "success", "output_path": output_pptx_path}), 200

@app.route('/download', methods=['GET'])
def download_file():
    file_path = request.args.get('file_path')
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({"error": "File not found"}), 404

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")
