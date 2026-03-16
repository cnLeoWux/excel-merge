"""
Excel Merge Tool - HTTP API Server
Author: Leo Wu leo.wux@lego.com
Date: 2026-03-04
Description: HTTP API service for Excel merge operations with file upload support
"""

import os
import tempfile
import uuid
from pathlib import Path
from datetime import datetime
from flask import Flask, request, send_file, jsonify, render_template_string
from werkzeug.utils import secure_filename
from utils import process_excel_files, write_result_file

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = Path("uploads")
RESULT_FOLDER = Path("results")
ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.csv'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

# Create directories
UPLOAD_FOLDER.mkdir(exist_ok=True)
RESULT_FOLDER.mkdir(exist_ok=True)


def allowed_file(filename):
    """Check if file extension is allowed"""
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Simple HTML test page"""
    html = '''
<!DOCTYPE html>
<html>
<head>
    <title>Excel Merge Tool API</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 50px auto; padding: 20px; }
        h1 { color: #333; }
        .upload-form { border: 2px dashed #ccc; padding: 30px; border-radius: 10px; text-align: center; }
        .file-input { margin: 15px 0; }
        button { background: #007bff; color: white; padding: 12px 30px; border: none; border-radius: 5px; cursor: pointer; }
        button:hover { background: #0056b3; }
        .info { background: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 20px; }
        code { background: #e9ecef; padding: 2px 6px; border-radius: 3px; }
    </style>
</head>
<body>
    <h1>📊 Excel Merge Tool</h1>
    
    <div class="upload-form">
        <h2>上传文件进行匹配</h2>
        <form action="/merge" method="post" enctype="multipart/form-data">
            <div class="file-input">
                <label>订单文件 (Order):</label><br>
                <input type="file" name="order_file" accept=".xlsx,.xls,.csv" required>
            </div>
            <div class="file-input">
                <label>支付流水文件 (Payment):</label><br>
                <input type="file" name="payment_file" accept=".xlsx,.xls,.csv" required>
            </div>
            <button type="submit">开始处理</button>
        </form>
    </div>
    
    <div class="info">
        <h3>API 使用说明</h3>
        <p><strong>Endpoint:</strong> <code>POST /merge</code></p>
        <p><strong>Parameters:</strong></p>
        <ul>
            <li><code>order_file</code> - 订单数据文件 (Excel/CSV)</li>
            <li><code>payment_file</code> - 支付流水文件 (Excel/CSV)</li>
        </ul>
        <p><strong>cURL 示例:</strong></p>
        <pre><code>curl -X POST http://localhost:5000/merge \\
  -F "order_file=@orders.xlsx" \\
  -F "payment_file=@payments.xlsx" \\
  --output result.xlsx</code></pre>
    </div>
</body>
</html>
    '''
    return render_template_string(html)


@app.route('/health')
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'service': 'excel-merge-api'
    })


@app.route('/merge', methods=['POST'])
def merge_files():
    """
    Merge two Excel files via HTTP POST
    
    Form Data:
        - order_file: Order data file (required)
        - payment_file: Payment/refund data file (required)
    
    Returns:
        - Processed Excel file as attachment
        - Or JSON error message
    """
    try:
        # Check if files are present
        if 'order_file' not in request.files:
            return jsonify({'error': 'Missing order_file'}), 400
        
        if 'payment_file' not in request.files:
            return jsonify({'error': 'Missing payment_file'}), 400
        
        order_file = request.files['order_file']
        payment_file = request.files['payment_file']
        
        # Validate files
        if order_file.filename == '':
            return jsonify({'error': 'No order file selected'}), 400
        
        if payment_file.filename == '':
            return jsonify({'error': 'No payment file selected'}), 400
        
        # Check file extensions
        if not allowed_file(order_file.filename):
            return jsonify({
                'error': f'Invalid order file type. Allowed: {", ".join(ALLOWED_EXTENSIONS)}'
            }), 400
        
        if not allowed_file(payment_file.filename):
            return jsonify({
                'error': f'Invalid payment file type. Allowed: {", ".join(ALLOWED_EXTENSIONS)}'
            }), 400
        
        # Generate unique session ID
        session_id = str(uuid.uuid4())[:8]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Save uploaded files
        order_filename = secure_filename(f"{session_id}_order_{order_file.filename}")
        payment_filename = secure_filename(f"{session_id}_payment_{payment_file.filename}")
        
        order_path = UPLOAD_FOLDER / order_filename
        payment_path = UPLOAD_FOLDER / payment_filename
        
        order_file.save(order_path)
        payment_file.save(payment_path)
        
        print(f"[{session_id}] Files saved: {order_path}, {payment_path}")
        
        # Process files
        print(f"[{session_id}] Starting merge process...")
        result_df = process_excel_files(str(order_path), str(payment_path), verbose=False)
        
        # Determine output filename
        original_ext = Path(order_file.filename).suffix
        result_filename = f"merged_result_{timestamp}_{session_id}{original_ext}"
        result_path = RESULT_FOLDER / result_filename
        
        # Save result
        write_result_file(result_df, result_path)
        print(f"[{session_id}] Result saved: {result_path}")
        
        # Clean up uploaded files (optional - comment out to keep them)
        # order_path.unlink()
        # payment_path.unlink()
        
        # Return the result file
        return send_file(
            result_path,
            as_attachment=True,
            download_name=f"merged_{order_file.filename}",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Error processing request: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'error': 'Processing failed',
            'message': str(e)
        }), 500


@app.route('/merge/json', methods=['POST'])
def merge_files_json():
    """
    Merge two Excel files and return JSON response with download URL
    Useful for async processing scenarios
    """
    try:
        if 'order_file' not in request.files or 'payment_file' not in request.files:
            return jsonify({'error': 'Missing required files'}), 400
        
        order_file = request.files['order_file']
        payment_file = request.files['payment_file']
        
        if order_file.filename == '' or payment_file.filename == '':
            return jsonify({'error': 'Empty filename'}), 400
        
        # Generate unique session ID
        session_id = str(uuid.uuid4())[:8]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Save files
        order_filename = secure_filename(f"{session_id}_order_{order_file.filename}")
        payment_filename = secure_filename(f"{session_id}_payment_{payment_file.filename}")
        
        order_path = UPLOAD_FOLDER / order_filename
        payment_path = UPLOAD_FOLDER / payment_filename
        
        order_file.save(order_path)
        payment_file.save(payment_path)
        
        # Process
        result_df = process_excel_files(str(order_path), str(payment_path), verbose=False)
        
        # Save result
        original_ext = Path(order_file.filename).suffix
        result_filename = f"merged_result_{timestamp}_{session_id}{original_ext}"
        result_path = RESULT_FOLDER / result_filename
        write_result_file(result_df, result_path)
        
        # Get statistics
        total_rows = len(result_df)
        matched_rows = result_df['支付手续费'].notna().sum()
        
        return jsonify({
            'success': True,
            'session_id': session_id,
            'download_url': f'/download/{result_filename}',
            'statistics': {
                'total_rows': int(total_rows),
                'matched_rows': int(matched_rows),
                'match_rate': f"{matched_rows/total_rows*100:.1f}%" if total_rows > 0 else "0%"
            },
            'files': {
                'order': order_file.filename,
                'payment': payment_file.filename,
                'result': result_filename
            }
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/download/<filename>')
def download_file(filename):
    """Download a processed result file"""
    file_path = RESULT_FOLDER / filename
    if not file_path.exists():
        return jsonify({'error': 'File not found'}), 404
    
    return send_file(
        file_path,
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    print("=" * 60)
    print("Excel Merge Tool - HTTP API Server")
    print("=" * 60)
    print(f"Upload folder: {UPLOAD_FOLDER.absolute()}")
    print(f"Results folder: {RESULT_FOLDER.absolute()}")
    print("\nEndpoints:")
    print("  GET  /          - Web interface")
    print("  GET  /health    - Health check")
    print("  POST /merge     - Upload and merge (returns file)")
    print("  POST /merge/json - Upload and merge (returns JSON)")
    print("  GET  /download/<file> - Download result")
    print("\nStarting server on http://localhost:5000")
    print("=" * 60)
    
    app.run(host='0.0.0.0', port=5000, debug=True)
