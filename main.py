from flask import Flask, request, jsonify, render_template_string
from flask_cors import CORS
import pandas as pd
import os
import re
import shutil
from pathlib import Path
import logging

app = Flask(__name__)
CORS(app)  # 解决跨域问题

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def clean_filename(filename):
    """清理文件名，移除不合法字符"""
    # 移除或替换Windows文件名中不允许的字符
    illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in illegal_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def extract_name_from_filename(filename):
    """从文件名中提取姓名"""
    # 移除文件扩展名
    name_part = os.path.splitext(filename)[0]

    # 记录原始文件名用于调试
    logger.info(f"正在提取姓名，原始文件名: {filename}")

    # 常见的提取姓名的模式 - 按优先级排序
    patterns = [
        # 1. 标准格式：学号+空格+姓名+空格
        r'^\d+\s+([^\s\d]+)\s+',

        # 2. 学号紧接姓名：202441050637 24级计算机工程专升本6班叶梓源37
        r'^\d+\s+.*?([^\d\s]{2,4})(?:\d+|第)',

        # 3. 学号+空格+班级信息+姓名：202441050610 24级计算机工程专升本6班 肖思进 10
        r'^\d+\s+.*?班\s+([^\s\d]+)\s+\d+',

        # 4. 学号+姓名（无空格）：202441050605吴碧瑜
        r'^\d+([^\d\s+]{2,4})',

        # 5. 姓名+学号格式
        r'^([^\d\s]{2,4})\s*\+?\s*\d+',

        # 6. 查找班级信息后的姓名
        r'班([^\d\s]{2,4})(?:\d+|\s|第)',

        # 7. 通用中文姓名匹配（2-4个字符）
        r'([^\d\s]{2,4})(?=\s|\d|第|实验|\+)',
    ]

    for i, pattern in enumerate(patterns):
        matches = re.findall(pattern, name_part)
        if matches:
            for match in matches:
                potential_name = match.strip()
                # 过滤掉明显不是姓名的内容
                if (len(potential_name) >= 2 and
                        not re.match(r'^\d+$', potential_name) and
                        potential_name not in ['级计算机工程专升本', '计算机工程专升本', '实验', '课程', '报告',
                                               '电子版', '云计算', '新版'] and
                        not any(word in potential_name for word in
                                ['级', '班', '实验', '报告', '课程', '计算机', '工程', '专升本'])):
                    logger.info(f"使用模式 {i + 1} 提取到姓名: {potential_name}")
                    return potential_name

    # 如果没有匹配到，尝试从文件名中找到所有中文字符组合
    chinese_groups = re.findall(r'[\u4e00-\u9fff]{2,4}', name_part)
    if chinese_groups:
        # 过滤掉常见的非姓名词汇
        exclude_words = ['级计算机工程专升本', '计算机工程专升本', '实验报告', '课程实验', '电子版', '云计算', '新版']
        for char_group in chinese_groups:
            if (2 <= len(char_group) <= 4 and
                    char_group not in exclude_words and
                    not any(word in char_group for word in ['级', '班', '实验', '报告', '课程', '计算机', '工程'])):
                logger.info(f"使用中文字符匹配提取到姓名: {char_group}")
                return char_group

    logger.warning(f"无法从文件名中提取姓名: {filename}")
    return None


def find_student_info(name, df):
    """根据姓名在Excel表格中查找学生信息"""
    # 清理输入姓名
    name = str(name).strip()

    # 精确匹配
    for idx, row in df.iterrows():
        student_name = str(row['姓名']).strip()
        if student_name == name:
            return student_name, str(row['学号']).strip()

    # 模糊匹配（包含关系）
    for idx, row in df.iterrows():
        student_name = str(row['姓名']).strip()
        if name in student_name or student_name in name:
            return student_name, str(row['学号']).strip()

    # 尝试部分匹配（去除可能的特殊字符）
    cleaned_name = re.sub(r'[^\u4e00-\u9fff]', '', name)  # 只保留中文字符
    if cleaned_name and cleaned_name != name:
        for idx, row in df.iterrows():
            student_name = str(row['姓名']).strip()
            cleaned_student_name = re.sub(r'[^\u4e00-\u9fff]', '', student_name)
            if cleaned_name == cleaned_student_name:
                return student_name, str(row['学号']).strip()

    return None, None


def rename_word_files(excel_path, word_folder, name_format):
    """重命名Word文档"""
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path)

        # 打印Excel文件的基本信息用于调试
        logger.info(f"Excel文件形状: {df.shape}")
        logger.info(f"Excel列名: {df.columns.tolist()}")

        # 获取姓名和学号数据
        # 从第1行（索引0）开始，A列和C列
        if len(df.columns) >= 3:
            df_students = df.iloc[0:50, [0, 2]].copy()
            df_students.columns = ['姓名', '学号']
        else:
            return {"success": False, "message": "Excel文件格式不正确，需要至少3列"}

        # 清理数据
        df_students = df_students.dropna()
        df_students['姓名'] = df_students['姓名'].astype(str).str.strip()
        df_students['学号'] = df_students['学号'].astype(str).str.strip()

        # 输出调试信息
        logger.info(f"提取的学生数据行数: {len(df_students)}")

        # 获取Word文件夹中的所有Word文档
        word_folder = Path(word_folder)
        if not word_folder.exists():
            return {"success": False, "message": f"Word文件夹不存在: {word_folder}"}

        word_files = []
        for ext in ['*.docx', '*.doc']:
            word_files.extend(word_folder.glob(ext))

        if not word_files:
            return {"success": False, "message": "未找到Word文档"}

        results = []
        success_count = 0

        for word_file in word_files:
            try:
                original_name = word_file.name

                # 从文件名中提取姓名
                extracted_name = extract_name_from_filename(original_name)

                if not extracted_name:
                    results.append({
                        "original": original_name,
                        "new": "未处理",
                        "status": "无法提取姓名"
                    })
                    continue

                # 在Excel中查找学生信息
                student_name, student_id = find_student_info(extracted_name, df_students)

                if not student_name or not student_id:
                    results.append({
                        "original": original_name,
                        "new": "未处理",
                        "status": f"未找到匹配的学生信息 (提取的姓名: {extracted_name})"
                    })
                    continue

                # 生成新的文件名
                target_filename = f"{student_id} {student_name} 实验（训）报告【电子版】-{name_format}{word_file.suffix}"
                target_filename = clean_filename(target_filename)
                target_path = word_folder / target_filename

                # 如果目标文件已存在且不是同一个文件，添加序号
                final_filename = target_filename
                final_path = target_path

                if target_path.exists() and target_path.samefile(word_file) == False:
                    counter = 1
                    while final_path.exists():
                        name_without_ext = f"{student_id} {student_name} 实验（训）报告【电子版】-{name_format}({counter})"
                        final_filename = clean_filename(name_without_ext) + word_file.suffix
                        final_path = word_folder / final_filename
                        counter += 1
                elif target_path.exists() and target_path.samefile(word_file):
                    results.append({
                        "original": original_name,
                        "new": "未处理",
                        "status": "文件名已经正确，无需修改"
                    })
                    continue

                # 执行重命名
                try:
                    word_file.rename(final_path)
                    success_count += 1

                    results.append({
                        "original": original_name,
                        "new": final_filename,
                        "status": "成功"
                    })
                except Exception as rename_error:
                    results.append({
                        "original": original_name,
                        "new": "未处理",
                        "status": f"重命名失败: {str(rename_error)}"
                    })
                    continue

            except Exception as e:
                results.append({
                    "original": word_file.name,
                    "new": "未处理",
                    "status": f"处理失败: {str(e)}"
                })

        return {
            "success": True,
            "message": f"处理完成！成功重命名 {success_count} 个文件",
            "results": results,
            "total": len(word_files),
            "success_count": success_count
        }

    except Exception as e:
        logger.error(f"处理过程中出现错误: {str(e)}")
        return {"success": False, "message": f"处理失败: {str(e)}"}


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/rename', methods=['POST'])
def rename_files():
    try:
        data = request.json
        excel_path = data.get('excel_path')
        word_folder = data.get('word_folder')
        name_format = data.get('name_format')

        if not all([excel_path, word_folder, name_format]):
            return jsonify({"success": False, "message": "请填写所有必需字段"})

        # 检查文件和文件夹是否存在
        if not os.path.exists(excel_path):
            return jsonify({"success": False, "message": f"Excel文件不存在: {excel_path}"})

        if not os.path.exists(word_folder):
            return jsonify({"success": False, "message": f"Word文件夹不存在: {word_folder}"})

        result = rename_word_files(excel_path, word_folder, name_format)
        return jsonify(result)

    except Exception as e:
        logger.error(f"API处理错误: {str(e)}")
        return jsonify({"success": False, "message": f"服务器错误: {str(e)}"})


# HTML模板
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Word文档重命名工具</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }

        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
        }

        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }

        .form-container {
            padding: 40px;
        }

        .form-group {
            margin-bottom: 25px;
        }

        .form-group label {
            display: block;
            margin-bottom: 8px;
            font-weight: bold;
            color: #333;
            font-size: 1.1em;
        }

        .form-group input[type="text"] {
            width: 100%;
            padding: 15px;
            border: 2px solid #e1e1e1;
            border-radius: 8px;
            font-size: 1em;
            transition: all 0.3s ease;
        }

        .form-group input[type="text"]:focus {
            outline: none;
            border-color: #4facfe;
            box-shadow: 0 0 0 3px rgba(79, 172, 254, 0.1);
        }

        .btn {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            padding: 15px 30px;
            border: none;
            border-radius: 8px;
            font-size: 1.1em;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.3s ease;
            width: 100%;
        }

        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(0,0,0,0.1);
        }

        .btn:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
        }

        .result {
            margin-top: 30px;
            padding: 20px;
            border-radius: 8px;
            display: none;
        }

        .result.success {
            background: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }

        .result.error {
            background: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }

        .result-details {
            max-height: 400px;
            overflow-y: auto;
            margin-top: 15px;
        }

        .result-item {
            padding: 8px 0;
            border-bottom: 1px solid rgba(0,0,0,0.1);
        }

        .result-item:last-child {
            border-bottom: none;
        }

        .loading {
            display: none;
            text-align: center;
            margin-top: 20px;
        }

        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #4facfe;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .help-text {
            font-size: 0.9em;
            color: #666;
            margin-top: 5px;
        }

        .example {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin-top: 10px;
            border-left: 4px solid #4facfe;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 Word文档重命名工具</h1>
            <p>根据Excel表格自动规范化Word文档命名</p>
        </div>

        <div class="form-container">
            <form id="renameForm">
                <div class="form-group">
                    <label for="nameFormat">📝 课程名称格式:</label>
                    <input type="text" id="nameFormat" name="nameFormat" placeholder="例如：云计算课程实验课程九(新版)" required>
                    <div class="help-text">这个名称将出现在重命名后的文件名中</div>
                    <div class="example">
                        <strong>示例：</strong>如果填写"云计算课程实验课程九(新版)"<br>
                        生成的文件名将是：202441050639 朱联桉 实验（训）报告【电子版】-云计算课程实验课程九(新版).docx
                    </div>
                </div>

                <div class="form-group">
                    <label for="excelPath">📊 Excel文件路径:</label>
                    <input type="text" id="excelPath" name="excelPath" placeholder="例如：C:\\Users\\用户名\\Desktop\\学生名单.xlsx" required>
                    <div class="help-text">包含学生姓名（A列）和学号（C列）的Excel文件完整路径</div>
                </div>

                <div class="form-group">
                    <label for="wordFolder">📁 Word文件夹路径:</label>
                    <input type="text" id="wordFolder" name="wordFolder" placeholder="例如：C:\\Users\\用户名\\Desktop\\作业文件夹" required>
                    <div class="help-text">包含需要重命名的Word文档的文件夹路径</div>
                </div>

                <button type="submit" class="btn" id="submitBtn">
                    🚀 开始重命名
                </button>
            </form>

            <div class="loading" id="loading">
                <div class="spinner"></div>
                <p>正在处理文件，请稍候...</p>
            </div>

            <div class="result" id="result">
                <div id="resultMessage"></div>
                <div class="result-details" id="resultDetails"></div>
            </div>
        </div>
    </div>

    <script>
        // 表单提交事件
        document.getElementById('renameForm').addEventListener('submit', async function(e) {
            e.preventDefault();

            const submitBtn = document.getElementById('submitBtn');
            const loading = document.getElementById('loading');
            const result = document.getElementById('result');
            const resultMessage = document.getElementById('resultMessage');
            const resultDetails = document.getElementById('resultDetails');

            // 获取表单数据
            const formData = {
                name_format: document.getElementById('nameFormat').value.trim(),
                excel_path: document.getElementById('excelPath').value.trim(),
                word_folder: document.getElementById('wordFolder').value.trim()
            };

            // 验证输入
            if (!formData.name_format || !formData.excel_path || !formData.word_folder) {
                alert('请填写所有必需字段！');
                return;
            }

            // 显示加载状态
            submitBtn.disabled = true;
            submitBtn.textContent = '处理中...';
            loading.style.display = 'block';
            result.style.display = 'none';

            try {
                const response = await fetch('/rename', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(formData)
                });

                const data = await response.json();

                // 显示结果
                result.style.display = 'block';
                result.className = 'result ' + (data.success ? 'success' : 'error');

                if (data.success) {
                    resultMessage.innerHTML = `
                        <h3>✅ ${data.message}</h3>
                        <p>总文件数: ${data.total} | 成功处理: ${data.success_count}</p>
                    `;

                    if (data.results && data.results.length > 0) {
                        let detailsHtml = '<h4>处理详情:</h4>';
                        data.results.forEach(item => {
                            const statusClass = item.status === '成功' ? 'success' : 'error';
                            detailsHtml += `
                                <div class="result-item">
                                    <strong>原文件名:</strong> ${item.original}<br>
                                    <strong>新文件名:</strong> ${item.new}<br>
                                    <strong>状态:</strong> <span class="${statusClass}">${item.status}</span>
                                </div>
                            `;
                        });
                        resultDetails.innerHTML = detailsHtml;
                    }
                } else {
                    resultMessage.innerHTML = `<h3>❌ 处理失败</h3><p>${data.message}</p>`;
                    resultDetails.innerHTML = '';
                }

            } catch (error) {
                result.style.display = 'block';
                result.className = 'result error';
                resultMessage.innerHTML = `<h3>❌ 网络错误</h3><p>请检查服务器连接: ${error.message}</p>`;
                resultDetails.innerHTML = '';
            }

            // 恢复按钮状态
            submitBtn.disabled = false;
            submitBtn.textContent = '🚀 开始重命名';
            loading.style.display = 'none';
        });
    </script>
</body>
</html>
'''

if __name__ == '__main__':
    print("🚀 Word文档重命名工具启动中...")
    print("📖 使用说明:")
    print("1. 准备包含学生姓名和学号的Excel文件")
    print("2. 将需要重命名的Word文档放在指定文件夹中")
    print("3. 在浏览器中打开 http://localhost:5002")
    print("4. 填写相关信息并开始处理")
    print("\n⚠️  注意: 请在运行前备份原始文件！")
    print("=" * 50)

    app.run(host='0.0.0.0', port=5002, debug=True)