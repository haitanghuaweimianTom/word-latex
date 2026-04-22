"""
Docx-LaTeX 转换工具 Web 服务
Flask 后端 + 前端页面
"""

from flask import Flask, request, jsonify, send_file, render_template_string, Response
import os
import sys
import tempfile
import io

# 添加转换模块路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'word2latex', 'src'))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'latex2word', 'src'))

from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# 允许的文件类型
ALLOWED_DOCX = {'docx'}
ALLOWED_TEX = {'tex'}


def allowed_file(filename, allowed_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set


@app.route('/')
def index():
    """主页"""
    return render_template_string(INDEX_HTML)


@app.route('/api/word2latex', methods=['POST'])
def word_to_latex():
    """Word → LaTeX 转换 API"""
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400

    if not allowed_file(file.filename, ALLOWED_DOCX):
        return jsonify({'error': '只支持 .docx 文件'}), 400

    try:
        # 保存上传的文件
        filename = secure_filename(file.filename)
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        temp_input.write(file.read())
        temp_input.close()

        # 执行转换
        from cli import main as word2latex_main

        # 重定向 stdout 到捕获输出
        old_stdout = sys.stdout
        sys.stdout = captured_output = io.StringIO()

        # 临时修改 sys.argv
        old_argv = sys.argv
        sys.argv = ['word2latex', temp_input.name, '--package-mode']

        try:
            word2latex_main()
            latex_code = captured_output.getvalue()
        finally:
            sys.stdout = old_stdout
            sys.argv = old_argv

        # 清理临时文件
        os.unlink(temp_input.name)

        return Response(
            latex_code,
            mimetype='text/plain',
            headers={'Content-Disposition': f'attachment;filename={filename.rsplit(".", 1)[0]}.tex'}
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/latex2word', methods=['POST'])
def latex_to_word():
    """LaTeX → Word 转换 API"""
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400

    if not allowed_file(file.filename, ALLOWED_TEX):
        return jsonify({'error': '只支持 .tex 文件'}), 400

    try:
        from cli import main as latex2word_main

        # 保存上传的文件
        filename = secure_filename(file.filename)
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.tex')
        temp_input.write(file.read())
        temp_input.close()

        # 创建临时输出文件
        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        temp_output.close()

        # 执行转换
        old_argv = sys.argv
        sys.argv = ['latex2word', temp_input.name, '-o', temp_output.name]

        try:
            latex2word_main()
        except SystemExit:
            pass
        finally:
            sys.argv = old_argv

        # 读取输出文件
        with open(temp_output.name, 'rb') as f:
            docx_data = f.read()

        # 清理临时文件
        os.unlink(temp_input.name)
        os.unlink(temp_output.name)

        return Response(
            docx_data,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers={'Content-Disposition': f'attachment;filename={filename.rsplit(".", 1)[0]}.docx'}
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


# 前端 HTML 模板
INDEX_HTML = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Docx-LaTeX 转换工具</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 40px 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
        }

        h1 {
            text-align: center;
            color: white;
            font-size: 2.5rem;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }

        .subtitle {
            text-align: center;
            color: rgba(255,255,255,0.9);
            margin-bottom: 40px;
            font-size: 1.1rem;
        }

        .cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 30px;
        }

        .card {
            background: white;
            border-radius: 20px;
            padding: 30px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }

        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 25px 70px rgba(0,0,0,0.35);
        }

        .card-header {
            display: flex;
            align-items: center;
            margin-bottom: 25px;
        }

        .card-icon {
            width: 60px;
            height: 60px;
            border-radius: 15px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.8rem;
            margin-right: 15px;
        }

        .card-icon.docx {
            background: linear-gradient(135deg, #2b5797, #1e3c5a);
        }

        .card-icon.latex {
            background: linear-gradient(135deg, #008080, #004d4d);
        }

        .card-title {
            font-size: 1.4rem;
            color: #333;
            font-weight: 600;
        }

        .card-desc {
            color: #666;
            font-size: 0.95rem;
            line-height: 1.6;
            margin-bottom: 25px;
        }

        .upload-area {
            border: 2px dashed #ddd;
            border-radius: 12px;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            position: relative;
        }

        .upload-area:hover {
            border-color: #667eea;
            background: rgba(102, 126, 234, 0.05);
        }

        .upload-area.dragover {
            border-color: #667eea;
            background: rgba(102, 126, 234, 0.1);
        }

        .upload-icon {
            font-size: 3rem;
            margin-bottom: 15px;
        }

        .upload-text {
            color: #666;
            font-size: 1rem;
        }

        .upload-text span {
            color: #667eea;
            font-weight: 600;
        }

        .file-input {
            position: absolute;
            width: 100%;
            height: 100%;
            top: 0;
            left: 0;
            opacity: 0;
            cursor: pointer;
        }

        .selected-file {
            display: none;
            margin-top: 15px;
            padding: 12px 15px;
            background: #f5f5f5;
            border-radius: 8px;
            font-size: 0.9rem;
            color: #333;
        }

        .selected-file.show {
            display: flex;
            align-items: center;
        }

        .selected-file-icon {
            margin-right: 10px;
            font-size: 1.2rem;
        }

        .selected-file-name {
            flex: 1;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .convert-btn {
            width: 100%;
            padding: 15px;
            margin-top: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 10px;
            font-size: 1.1rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
        }

        .convert-btn:hover {
            transform: scale(1.02);
            box-shadow: 0 10px 30px rgba(102, 126, 234, 0.4);
        }

        .convert-btn:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
        }

        .convert-btn .spinner {
            display: none;
            width: 20px;
            height: 20px;
            border: 2px solid rgba(255,255,255,0.3);
            border-top-color: white;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
        }

        .convert-btn.loading .spinner {
            display: block;
        }

        .convert-btn.loading .btn-text {
            display: none;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .arrow {
            display: flex;
            justify-content: center;
            align-items: center;
            margin: 30px 0;
            color: white;
            font-size: 2rem;
        }

        .feature-list {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 20px;
        }

        .feature-tag {
            background: #f0f0f0;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.8rem;
            color: #555;
        }

        .card.word2latex .convert-btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }

        .card.latex2word .convert-btn {
            background: linear-gradient(135deg, #008080 0%, #004d4d 100%);
        }

        .card.latex2word .convert-btn:hover {
            box-shadow: 0 10px 30px rgba(0, 128, 128, 0.4);
        }

        footer {
            text-align: center;
            color: rgba(255,255,255,0.7);
            margin-top: 50px;
            font-size: 0.9rem;
        }

        @media (max-width: 850px) {
            .cards {
                grid-template-columns: 1fr;
            }

            .arrow {
                transform: rotate(90deg);
                margin: 20px 0;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Docx-LaTeX 转换工具</h1>
        <p class="subtitle">轻松实现 Word 与 LaTeX 文档的相互转换</p>

        <div class="cards">
            <!-- Word 转 LaTeX -->
            <div class="card word2latex">
                <div class="card-header">
                    <div class="card-icon docx">W</div>
                    <div class="card-title">Word → LaTeX</div>
                </div>
                <p class="card-desc">
                    将 Word 文档 (.docx) 转换为 LaTeX 代码，
                    支持数学公式、表格、标题等元素的精确转换。
                </p>
                <div class="upload-area" id="upload-word">
                    <div class="upload-icon">📄</div>
                    <p class="upload-text">
                        拖拽或<span>点击选择</span> Word 文件<br>
                        <small>支持 .docx 格式</small>
                    </p>
                    <input type="file" class="file-input" accept=".docx" id="word-file">
                </div>
                <div class="selected-file" id="word-selected">
                    <span class="selected-file-icon">📄</span>
                    <span class="selected-file-name" id="word-filename">文件.docx</span>
                </div>
                <button class="convert-btn" id="btn-word" disabled>
                    <span class="spinner"></span>
                    <span class="btn-text">转换为 LaTeX</span>
                </button>
                <div class="feature-list">
                    <span class="feature-tag">✅ 标题转换</span>
                    <span class="feature-tag">✅ 数学公式</span>
                    <span class="feature-tag">✅ 表格支持</span>
                    <span class="feature-tag">✅ 加粗斜体</span>
                    <span class="feature-tag">✅ 图片提取</span>
                </div>
            </div>

            <!-- LaTeX 转 Word -->
            <div class="card latex2word">
                <div class="card-header">
                    <div class="card-icon latex">L</div>
                    <div class="card-title">LaTeX → Word</div>
                </div>
                <p class="card-desc">
                    将 LaTeX 文件 (.tex) 转换为 Word 文档 (.docx)，
                    支持多种 LaTeX 元素的智能转换。
                </p>
                <div class="upload-area" id="upload-latex">
                    <div class="upload-icon">📝</div>
                    <p class="upload-text">
                        拖拽或<span>点击选择</span> LaTeX 文件<br>
                        <small>支持 .tex 格式</small>
                    </p>
                    <input type="file" class="file-input" accept=".tex" id="latex-file">
                </div>
                <div class="selected-file" id="latex-selected">
                    <span class="selected-file-icon">📝</span>
                    <span class="selected-file-name" id="latex-filename">文件.tex</span>
                </div>
                <button class="convert-btn" id="btn-latex" disabled>
                    <span class="spinner"></span>
                    <span class="btn-text">转换为 Word</span>
                </button>
                <div class="feature-list">
                    <span class="feature-tag">✅ 标题转换</span>
                    <span class="feature-tag">✅ 公式转换</span>
                    <span class="feature-tag">✅ 表格支持</span>
                    <span class="feature-tag">✅ 列表转换</span>
                    <span class="feature-tag">✅ 样式保留</span>
                </div>
            </div>
        </div>

        <footer>
            <p>Docx-LaTeX 转换工具 · 支持双向转换 · 安全本地处理</p>
        </footer>
    </div>

    <script>
        // Word → LaTeX
        const wordFileInput = document.getElementById('word-file');
        const wordSelected = document.getElementById('word-selected');
        const wordFilename = document.getElementById('word-filename');
        const wordBtn = document.getElementById('btn-word');

        wordFileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                wordFilename.textContent = file.name;
                wordSelected.classList.add('show');
                wordBtn.disabled = false;
            }
        });

        wordBtn.addEventListener('click', async () => {
            const file = wordFileInput.files[0];
            if (!file) return;

            wordBtn.classList.add('loading');
            wordBtn.disabled = true;

            const formData = new FormData();
            formData.append('file', file);

            try {
                const response = await fetch('/api/word2latex', {
                    method: 'POST',
                    body: formData
                });

                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = file.name.replace('.docx', '.tex');
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);
                } else {
                    const error = await response.json();
                    alert('转换失败: ' + error.error);
                }
            } catch (err) {
                alert('转换出错: ' + err.message);
            } finally {
                wordBtn.classList.remove('loading');
                wordBtn.disabled = false;
            }
        });

        // LaTeX → Word
        const latexFileInput = document.getElementById('latex-file');
        const latexSelected = document.getElementById('latex-selected');
        const latexFilename = document.getElementById('latex-filename');
        const latexBtn = document.getElementById('btn-latex');

        latexFileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                latexFilename.textContent = file.name;
                latexSelected.classList.add('show');
                latexBtn.disabled = false;
            }
        });

        latexBtn.addEventListener('click', async () => {
            const file = latexFileInput.files[0];
            if (!file) return;

            latexBtn.classList.add('loading');
            latexBtn.disabled = true;

            const formData = new FormData();
            formData.append('file', file);

            try {
                const response = await fetch('/api/latex2word', {
                    method: 'POST',
                    body: formData
                });

                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = file.name.replace('.tex', '.docx');
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);
                } else {
                    const error = await response.json();
                    alert('转换失败: ' + error.error);
                }
            } catch (err) {
                alert('转换出错: ' + err.message);
            } finally {
                latexBtn.classList.remove('loading');
                latexBtn.disabled = false;
            }
        });

        // 拖拽支持
        ['word', 'latex'].forEach(type => {
            const uploadArea = document.getElementById('upload-' + type);
            const fileInput = document.getElementById(type + '-file');

            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                uploadArea.addEventListener(eventName, (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                });
            });

            ['dragenter', 'dragover'].forEach(eventName => {
                uploadArea.addEventListener(eventName, () => {
                    uploadArea.classList.add('dragover');
                });
            });

            ['dragleave', 'drop'].forEach(eventName => {
                uploadArea.addEventListener(eventName, () => {
                    uploadArea.classList.remove('dragover');
                });
            });

            uploadArea.addEventListener('drop', (e) => {
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    fileInput.files = files;
                    fileInput.dispatchEvent(new Event('change'));
                }
            });
        });
    </script>
</body>
</html>
'''


if __name__ == '__main__':
    print("=" * 50)
    print("Docx-LaTeX 转换工具 Web 服务")
    print("=" * 50)
    print("服务地址: http://127.0.0.1:5000")
    print("按 Ctrl+C 停止服务")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=True)