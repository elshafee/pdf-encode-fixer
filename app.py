#!/usr/bin/env python3
"""
app.py - Modern Web Application Interface for PDF Arabic Text Fixer & Mark List Re-engineering.
"""

import os
import sys
import io
import base64
import json
import tempfile
from flask import Flask, render_template, request, jsonify, send_file
import fitz

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import arabic_fixer_core

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'))

# On Vercel / serverless platforms, filesystem is read-only except for /tmp
if os.environ.get('VERCEL') or not os.access(BASE_DIR, os.W_OK):
    OUTPUT_DIR = os.path.join(tempfile.gettempdir(), 'pdf_fixer_output')
else:
    OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

SAMPLES_DIR = os.path.join(BASE_DIR, 'pdf to test')
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/pdf_arabic_fixer.html')
def serve_standalone():
    """Serves the standalone client-side HTML fixer."""
    html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pdf_arabic_fixer.html')
    return send_file(html_path)

@app.route('/api/samples', methods=['GET'])
def get_samples():
    """Lists available test PDFs in 'pdf to test'."""
    samples = []
    if os.path.exists(SAMPLES_DIR):
        for f in sorted(os.listdir(SAMPLES_DIR)):
            if f.endswith('.pdf'):
                samples.append(f)
    return jsonify({"samples": samples})

@app.route('/api/process-sample/<sample_name>', methods=['POST'])
def process_sample(sample_name):
    """Processes a test sample PDF by name."""
    sample_path = os.path.join(SAMPLES_DIR, sample_name)
    if not os.path.exists(sample_path):
        return jsonify({"error": "Sample file not found"}), 404
    
    try:
        data = arabic_fixer_core.extract_mark_list_data(sample_path)
        
        # Generate initial redesigned PDF & preview
        base_name = sample_name.replace('.pdf', '')
        out_pdf = os.path.join(OUTPUT_DIR, f"{base_name}_Redesigned.pdf")
        arabic_fixer_core.generate_redesigned_pdf(data, out_pdf)
        
        # Also generate Excel
        out_xlsx = os.path.join(OUTPUT_DIR, f"{base_name}.xlsx")
        arabic_fixer_core.export_to_excel(data, out_xlsx)
        
        # Render preview image as base64
        doc = fitz.open(out_pdf)
        pix = doc[0].get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        b64_img = base64.b64encode(img_bytes).decode('utf-8')
        
        return jsonify({
            "success": True,
            "filename": sample_name,
            "data": data,
            "preview_image": f"data:image/png;base64,{b64_img}"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Uploads and processes any mark list PDF."""
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid .pdf file"}), 400
    
    try:
        pdf_bytes = file.read()
        data = arabic_fixer_core.extract_mark_list_data(pdf_bytes)
        
        base_name = os.path.splitext(file.filename)[0]
        out_pdf = os.path.join(OUTPUT_DIR, f"{base_name}_Redesigned.pdf")
        arabic_fixer_core.generate_redesigned_pdf(data, out_pdf)
        
        out_xlsx = os.path.join(OUTPUT_DIR, f"{base_name}.xlsx")
        arabic_fixer_core.export_to_excel(data, out_xlsx)
        
        doc = fitz.open(out_pdf)
        pix = doc[0].get_pixmap(dpi=150)
        img_bytes = pix.tobytes("png")
        b64_img = base64.b64encode(img_bytes).decode('utf-8')
        
        return jsonify({
            "success": True,
            "filename": file.filename,
            "data": data,
            "preview_image": f"data:image/png;base64,{b64_img}"
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/download-pdf', methods=['POST'])
def download_pdf():
    """Generates PDF from potentially edited JSON data and sends as download."""
    try:
        req_json = request.get_json(force=True)
        data = req_json.get('data')
        filename = req_json.get('filename', 'mark_list')
        base_name = os.path.splitext(filename)[0]
        
        out_pdf = os.path.join(OUTPUT_DIR, f"{base_name}_Redesigned.pdf")
        arabic_fixer_core.generate_redesigned_pdf(data, out_pdf)
        
        return send_file(
            out_pdf,
            as_attachment=True,
            download_name=f"{base_name}_Redesigned.pdf",
            mimetype='application/pdf'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/download-excel', methods=['POST'])
def download_excel():
    """Generates Excel workbook from JSON data and sends as download."""
    try:
        req_json = request.get_json(force=True)
        data = req_json.get('data')
        filename = req_json.get('filename', 'mark_list')
        base_name = os.path.splitext(filename)[0]
        
        out_xlsx = os.path.join(OUTPUT_DIR, f"{base_name}.xlsx")
        arabic_fixer_core.export_to_excel(data, out_xlsx)
        
        return send_file(
            out_xlsx,
            as_attachment=True,
            download_name=f"{base_name}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/batch-all', methods=['POST'])
def batch_all():
    """Processes all files in 'pdf to test' and generates outputs."""
    results = []
    if os.path.exists(SAMPLES_DIR):
        for fname in sorted(os.listdir(SAMPLES_DIR)):
            if not fname.endswith('.pdf'):
                continue
            fpath = os.path.join(SAMPLES_DIR, fname)
            try:
                data = arabic_fixer_core.extract_mark_list_data(fpath)
                base = fname.replace('.pdf', '')
                out_pdf = os.path.join(OUTPUT_DIR, f"{base}_Redesigned.pdf")
                out_xlsx = os.path.join(OUTPUT_DIR, f"{base}.xlsx")
                arabic_fixer_core.generate_redesigned_pdf(data, out_pdf)
                arabic_fixer_core.export_to_excel(data, out_xlsx)
                results.append({
                    "filename": fname,
                    "students": data["summary"]["count"],
                    "status": "success",
                    "pdf": f"{base}_Redesigned.pdf",
                    "excel": f"{base}.xlsx"
                })
            except Exception as e:
                results.append({"filename": fname, "status": "error", "error": str(e)})
    return jsonify({"success": True, "results": results})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    print(f"🚀 PDF Arabic Fixer Web Server running at http://127.0.0.1:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
