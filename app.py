# -*- coding: utf-8 -*-
"""
周报助手 - Flask版（用于线上部署）
"""
import os
import sys
import json
import re
import uuid
import threading
import time
from flask import Flask, request, send_file, jsonify, Response

app = Flask(__name__)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from ocr_parser import OCRParser, get_global_ocr
from excel_filler import ExcelFiller

# 临时目录（线上环境用 /tmp，本地用 temp）
TEMP_DIR = os.environ.get('TEMP_DIR', os.path.join(SCRIPT_DIR, 'temp'))
os.makedirs(TEMP_DIR, exist_ok=True)

# 应用启动时预加载OCR模型（避免第一次请求超时）
print('正在预加载OCR模型...', flush=True)
try:
    get_global_ocr()
    print('OCR模型加载完成', flush=True)
except Exception as e:
    print(f'OCR模型预加载失败: {e}', flush=True)


@app.route('/')
def index():
    return Response(HTML_PAGE, mimetype='text/html; charset=utf-8')


@app.route('/status')
def status():
    return jsonify({'status': 'ok'})


@app.route('/process', methods=['POST'])
def process():
    try:
        # 获取人员姓名
        target_name = request.form.get('target_name', '陈子怡').strip() or '陈子怡'

        # 保存模板文件
        template_file = request.files.get('template')
        if not template_file:
            return jsonify({'error': '请上传周报模板'}), 400

        template_path = os.path.join(TEMP_DIR, f'template_{uuid.uuid4().hex}.xlsx')
        template_file.save(template_path)

        # 保存图片文件
        image_paths = []
        for key in request.files:
            if key.startswith('image_'):
                f = request.files[key]
                if f and f.filename:
                    ext = os.path.splitext(f.filename)[1] or '.jpg'
                    img_path = os.path.join(TEMP_DIR, f'img_{uuid.uuid4().hex}{ext}')
                    f.save(img_path)
                    image_paths.append(img_path)

        if not image_paths:
            return jsonify({'error': '请上传至少一张图片'}), 400

        # 加载配置
        config_path = os.path.join(SCRIPT_DIR, 'config.json')
        text_corrections = {}
        team_info = ''
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            text_corrections = cfg.get('text_corrections', {})
            team_info = cfg.get('team_info', '')

        # OCR识别
        parser = OCRParser(target_name=target_name, text_corrections=text_corrections)
        parsed_results = []
        for img_path in image_paths:
            result = parser.parse_image(img_path)
            parsed_results.append(result)

        # 按日期排序
        def date_key(r):
            d = r.get('date')
            if not d:
                deadline = r.get('deadline') or ''
                m = re.search(r'(\d{1,2})月(\d{1,2})日', str(deadline))
                if m:
                    return (2026, int(m.group(1)), int(m.group(2)))
                return (9999, 99, 99)
            d = str(d)
            m = re.search(r'(\d{4})/(\d{1,2})/(\d{1,2})', d)
            if m:
                return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
            m = re.search(r'(\d{1,2})月(\d{1,2})日', d)
            if m:
                return (2026, int(m.group(1)), int(m.group(2)))
            return (9999, 99, 99)
        parsed_results.sort(key=date_key)

        # 生成输出文件
        output_filename = f'report_{uuid.uuid4().hex[:8]}.xlsx'
        output_path = os.path.join(TEMP_DIR, output_filename)

        filler = ExcelFiller(
            template_path=template_path,
            output_path=output_path,
            team_info=team_info
        )
        filler.fill(parsed_results)

        # 整理识别结果摘要
        summary = []
        for r in parsed_results:
            if r.get('found'):
                summary.append({
                    'date': r.get('date'),
                    'weekday': r.get('weekday'),
                    'count': len(r.get('work_items', [])),
                    'items': r.get('work_items', [])
                })
            else:
                summary.append({
                    'date': r.get('date') or '未知',
                    'weekday': r.get('weekday') or '',
                    'count': 0,
                    'items': [],
                    'error': r.get('reason', '')
                })

        # 延迟清理临时文件
        def cleanup():
            time.sleep(120)
            for p in [template_path] + image_paths + [output_path]:
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass
        threading.Thread(target=cleanup, daemon=True).start()

        return jsonify({
            'success': True,
            'download_url': f'/download/{output_filename}',
            'filename': output_filename,
            'summary': summary
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'处理失败：{str(e)}'}), 500


@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(TEMP_DIR, filename)
    if os.path.exists(filepath):
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return jsonify({'error': '文件不存在'}), 404


HTML_PAGE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>周报助手</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
}
.container {
    background: #fff;
    border-radius: 16px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    width: 100%;
    max-width: 680px;
    padding: 40px;
}
h1 {
    font-size: 28px;
    color: #333;
    margin-bottom: 8px;
    text-align: center;
}
.subtitle {
    text-align: center;
    color: #888;
    margin-bottom: 32px;
    font-size: 14px;
}
.form-group {
    margin-bottom: 24px;
}
label {
    display: block;
    font-weight: 600;
    color: #555;
    margin-bottom: 8px;
    font-size: 14px;
}
.file-input-wrapper {
    position: relative;
    border: 2px dashed #ddd;
    border-radius: 10px;
    padding: 24px;
    text-align: center;
    cursor: pointer;
    transition: all 0.3s;
    background: #fafafa;
}
.file-input-wrapper:hover {
    border-color: #667eea;
    background: #f0f0ff;
}
.file-input-wrapper.dragover {
    border-color: #667eea;
    background: #e8e8ff;
}
.file-input-wrapper input[type=file] {
    position: absolute;
    top: 0; left: 0;
    width: 100%; height: 100%;
    opacity: 0;
    cursor: pointer;
}
.file-icon {
    font-size: 36px;
    margin-bottom: 8px;
}
.file-text {
    color: #666;
    font-size: 14px;
}
.file-list {
    margin-top: 12px;
    font-size: 13px;
    color: #667eea;
}
.file-list span {
    display: inline-block;
    background: #f0f0ff;
    padding: 2px 10px;
    border-radius: 12px;
    margin: 2px;
}
input[type=text] {
    width: 100%;
    padding: 12px 16px;
    border: 1px solid #ddd;
    border-radius: 8px;
    font-size: 15px;
    transition: border-color 0.3s;
}
input[type=text]:focus {
    outline: none;
    border-color: #667eea;
}
.btn {
    width: 100%;
    padding: 14px;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff;
    border: none;
    border-radius: 10px;
    font-size: 16px;
    font-weight: 600;
    cursor: pointer;
    transition: transform 0.2s, opacity 0.3s;
}
.btn:hover { transform: translateY(-2px); }
.btn:disabled {
    opacity: 0.6;
    cursor: not-allowed;
    transform: none;
}
.progress-container {
    display: none;
    margin-top: 24px;
}
.progress-bar {
    width: 100%;
    height: 8px;
    background: #eee;
    border-radius: 4px;
    overflow: hidden;
}
.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, #667eea, #764ba2);
    width: 0%;
    transition: width 0.3s;
}
.progress-text {
    text-align: center;
    margin-top: 8px;
    font-size: 13px;
    color: #888;
}
.result {
    display: none;
    margin-top: 24px;
    padding: 20px;
    background: #f0fff4;
    border: 1px solid #c6f6d5;
    border-radius: 10px;
}
.result h3 {
    color: #22543d;
    margin-bottom: 12px;
}
.result-item {
    font-size: 13px;
    color: #555;
    margin-bottom: 6px;
    padding-left: 16px;
}
.download-btn {
    display: inline-block;
    margin-top: 12px;
    padding: 10px 24px;
    background: #38a169;
    color: #fff;
    text-decoration: none;
    border-radius: 8px;
    font-weight: 600;
}
.error {
    display: none;
    margin-top: 24px;
    padding: 16px;
    background: #fff5f5;
    border: 1px solid #feb2b2;
    border-radius: 10px;
    color: #c53030;
    font-size: 14px;
}
</style>
</head>
<body>
<div class="container">
    <h1>周报助手</h1>
    <p class="subtitle">自动识别排期图片中的工作内容，一键填入周报模板</p>

    <div class="form-group">
        <label>周报模板（Excel）</label>
        <div class="file-input-wrapper" id="templateWrapper">
            <input type="file" id="templateInput" accept=".xlsx,.xls">
            <div class="file-icon">📊</div>
            <div class="file-text">点击选择或拖拽Excel模板到此处</div>
            <div class="file-list" id="templateList"></div>
        </div>
    </div>

    <div class="form-group">
        <label>排期图片（可多选）</label>
        <div class="file-input-wrapper" id="imageWrapper">
            <input type="file" id="imageInput" accept="image/*" multiple>
            <div class="file-icon">🖼️</div>
            <div class="file-text">点击选择或拖拽图片到此处（支持多选）</div>
            <div class="file-list" id="imageList"></div>
        </div>
    </div>

    <div class="form-group">
        <label>提取人员姓名</label>
        <input type="text" id="targetName" value="陈子怡" placeholder="请输入要提取的人员姓名">
    </div>

    <button class="btn" id="submitBtn" onclick="submitForm()">开始生成周报</button>

    <div class="progress-container" id="progressContainer">
        <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
        <div class="progress-text" id="progressText">正在处理...</div>
    </div>

    <div class="result" id="resultBox">
        <h3>✅ 生成成功！</h3>
        <div id="resultSummary"></div>
        <a href="#" class="download-btn" id="downloadLink" download>📥 下载周报Excel</a>
    </div>

    <div class="error" id="errorBox"></div>
</div>

<script>
let templateFile = null;
let imageFiles = [];

document.getElementById('templateInput').addEventListener('change', function(e) {
    if (e.target.files.length > 0) {
        templateFile = e.target.files[0];
        document.getElementById('templateList').innerHTML = '<span>' + templateFile.name + '</span>';
    }
});

document.getElementById('imageInput').addEventListener('change', function(e) {
    imageFiles = Array.from(e.target.files);
    updateImageList();
});

['templateWrapper', 'imageWrapper'].forEach(function(id) {
    var wrapper = document.getElementById(id);
    var input = wrapper.querySelector('input[type=file]');
    wrapper.addEventListener('dragover', function(e) {
        e.preventDefault();
        wrapper.classList.add('dragover');
    });
    wrapper.addEventListener('dragleave', function() {
        wrapper.classList.remove('dragover');
    });
    wrapper.addEventListener('drop', function(e) {
        e.preventDefault();
        wrapper.classList.remove('dragover');
        input.files = e.dataTransfer.files;
        input.dispatchEvent(new Event('change'));
    });
});

function updateImageList() {
    var list = document.getElementById('imageList');
    if (imageFiles.length === 0) {
        list.innerHTML = '';
    } else {
        list.innerHTML = imageFiles.map(function(f) {
            return '<span>' + f.name + '</span>';
        }).join('');
    }
}

function submitForm() {
    if (!templateFile) {
        showError('请先选择周报模板');
        return;
    }
    if (imageFiles.length === 0) {
        showError('请至少选择一张图片');
        return;
    }

    var btn = document.getElementById('submitBtn');
    btn.disabled = true;
    btn.textContent = '处理中...';
    document.getElementById('resultBox').style.display = 'none';
    document.getElementById('errorBox').style.display = 'none';
    document.getElementById('progressContainer').style.display = 'block';
    setProgress(10, '正在上传文件...');

    var formData = new FormData();
    formData.append('template', templateFile);
    formData.append('target_name', document.getElementById('targetName').value || '陈子怡');
    imageFiles.forEach(function(f, i) {
        formData.append('image_' + i, f);
    });

    setProgress(30, '正在OCR识别图片内容...');

    fetch('/process', {
        method: 'POST',
        body: formData
    })
    .then(function(response) { return response.json(); })
    .then(function(data) {
        if (data.success) {
            setProgress(100, '完成！');
            showResult(data);
        } else {
            showError(data.error || '处理失败');
        }
    })
    .catch(function(err) {
        showError('网络错误：' + err.message);
    })
    .finally(function() {
        btn.disabled = false;
        btn.textContent = '开始生成周报';
        setTimeout(function() {
            document.getElementById('progressContainer').style.display = 'none';
        }, 1500);
    });
}

function setProgress(percent, text) {
    document.getElementById('progressFill').style.width = percent + '%';
    document.getElementById('progressText').textContent = text;
}

function showResult(data) {
    var box = document.getElementById('resultBox');
    var summary = document.getElementById('resultSummary');
    var html = '';
    data.summary.forEach(function(item) {
        if (item.count > 0) {
            html += '<div class="result-item">📅 ' + (item.date || '未知日期') + ' (' + (item.weekday || '') + ')：' + item.count + '项工作</div>';
        } else {
            html += '<div class="result-item">📅 ' + (item.date || '未知日期') + '：未找到（' + (item.error || '无数据') + '）</div>';
        }
    });
    summary.innerHTML = html;
    document.getElementById('downloadLink').href = data.download_url;
    document.getElementById('downloadLink').download = data.filename;
    box.style.display = 'block';
}

function showError(msg) {
    var box = document.getElementById('errorBox');
    box.textContent = '❌ ' + msg;
    box.style.display = 'block';
}
</script>
</body>
</html>'''


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8765))
    app.run(host='0.0.0.0', port=port, debug=False)
