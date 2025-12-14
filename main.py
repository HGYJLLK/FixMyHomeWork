from flask import Flask, request, jsonify, render_template_string
from flask_cors import CORS
import pandas as pd
import os
import re
import shutil
from pathlib import Path
import logging

app = Flask(__name__)
CORS(app)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def parse_excel_range(range_str):
    """解析Excel范围字符串，如 'A2:C12' 或 'A2-C12'"""
    range_str = range_str.replace('-', ':').upper()

    if ':' in range_str:
        start, end = range_str.split(':')

        # 解析列字母和行号
        start_col = re.findall(r'[A-Z]+', start)[0]
        start_row = int(re.findall(r'\d+', start)[0]) - 1  # pandas索引从0开始

        end_col = re.findall(r'[A-Z]+', end)[0]
        end_row = int(re.findall(r'\d+', end)[0]) - 1

        # 转换列字母为数字索引
        def col_to_num(col):
            num = 0
            for c in col:
                num = num * 26 + (ord(c) - ord('A')) + 1
            return num - 1

        start_col_idx = col_to_num(start_col)
        end_col_idx = col_to_num(end_col)

        return start_row, end_row, start_col_idx, end_col_idx
    else:
        raise ValueError("范围格式不正确，请使用 'A2:C12' 或 'A2-C12' 格式")


def load_excel_data(excel_path, name_range, id_range=None):
    """加载Excel数据"""
    df = pd.read_excel(excel_path, header=None)

    # 解析姓名范围
    name_start_row, name_end_row, name_start_col, name_end_col = parse_excel_range(name_range)

    # 提取姓名数据
    names_data = df.iloc[name_start_row:name_end_row + 1, name_start_col:name_end_col + 1]

    # 如果指定了学号范围
    if id_range:
        id_start_row, id_end_row, id_start_col, id_end_col = parse_excel_range(id_range)
        ids_data = df.iloc[id_start_row:id_end_row + 1, id_start_col:id_end_col + 1]

        # 合并数据
        student_data = []
        for i in range(len(names_data)):
            name = str(names_data.iloc[i, 0]).strip() if i < len(names_data) else ""
            student_id = str(ids_data.iloc[i, 0]).strip() if i < len(ids_data) else ""
            if name and name != 'nan' and student_id and student_id != 'nan':
                student_data.append({'name': name, 'id': student_id})
    else:
        # 只有姓名数据
        student_data = []
        for i in range(len(names_data)):
            name = str(names_data.iloc[i, 0]).strip()
            if name and name != 'nan':
                student_data.append({'name': name, 'id': ''})

    return student_data


def extract_info_from_filename(filename):
    """从文件名中提取各种信息"""
    name_part = os.path.splitext(filename)[0]
    logger.info(f"分析文件名: {filename}")

    extracted_info = {
        'names': [],
        'ids': [],
        'numbers': []
    }

    # 提取中文姓名（2-4个字符）
    chinese_names = re.findall(r'[\u4e00-\u9fff]{2,4}', name_part)
    for name in chinese_names:
        if not any(word in name for word in ['级', '班', '实验', '报告', '课程', '计算机', '工程', '电子版']):
            extracted_info['names'].append(name)

    # 提取数字（可能的学号）
    numbers = re.findall(r'\d{8,12}', name_part)  # 8-12位数字
    extracted_info['ids'].extend(numbers)

    # 提取所有数字
    all_numbers = re.findall(r'\d+', name_part)
    extracted_info['numbers'].extend(all_numbers)

    return extracted_info


def match_student_info(extracted_info, student_data, match_strategy):
    """匹配学生信息"""

    for strategy in match_strategy:
        if strategy == 'exact_name':
            # 精确姓名匹配
            for name in extracted_info['names']:
                for student in student_data:
                    if student['name'] == name:
                        return student, f"精确姓名匹配: {name}"

        elif strategy == 'partial_name':
            # 部分姓名匹配
            for name in extracted_info['names']:
                for student in student_data:
                    if name in student['name'] or student['name'] in name:
                        return student, f"部分姓名匹配: {name} -> {student['name']}"

        elif strategy == 'split_name':
            # 拆分姓名匹配（姓+名）
            for name in extracted_info['names']:
                if len(name) >= 2:
                    surname = name[0]  # 姓
                    given_name = name[1:]  # 名

                    for student in student_data:
                        if student['name'].startswith(surname) and given_name in student['name']:
                            return student, f"拆分姓名匹配: {surname}+{given_name} -> {student['name']}"

        elif strategy == 'exact_id':
            # 精确学号匹配
            for student_id in extracted_info['ids']:
                for student in student_data:
                    if student['id'] == student_id:
                        return student, f"精确学号匹配: {student_id}"

        elif strategy == 'partial_id':
            # 部分学号匹配
            for student_id in extracted_info['ids']:
                for student in student_data:
                    if student_id in student['id'] or student['id'] in student_id:
                        return student, f"部分学号匹配: {student_id} -> {student['id']}"

    return None, "未找到匹配"


def clean_filename(filename):
    """清理文件名，移除不合法字符"""
    illegal_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for char in illegal_chars:
        filename = filename.replace(char, '_')
    return filename.strip()


def rename_word_files(excel_path, word_folder, name_format, name_range, id_range, match_strategies):
    """重命名Word文档、PDF文件和图片文件"""
    try:
        # 加载Excel数据
        student_data = load_excel_data(excel_path, name_range, id_range)
        logger.info(f"加载学生数据: {len(student_data)} 条记录")

        # 获取文件夹
        word_folder = Path(word_folder)
        if not word_folder.exists():
            return {"success": False, "message": f"文件夹不存在: {word_folder}"}

        # 获取Word文件、PDF文件和图片文件
        word_files = []
        # Word文档
        for ext in ['*.docx', '*.doc']:
            word_files.extend(word_folder.glob(ext))
        # PDF文件
        for ext in ['*.pdf', '*.PDF']:
            word_files.extend(word_folder.glob(ext))
        # 图片文件
        for ext in ['*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp', '*.webp', '*.JPG', '*.JPEG', '*.PNG', '*.GIF', '*.BMP', '*.WEBP']:
            word_files.extend(word_folder.glob(ext))

        if not word_files:
            return {"success": False, "message": "未找到Word文档、PDF文件或图片文件"}

        results = []
        success_count = 0

        for word_file in word_files:
            try:
                original_name = word_file.name

                # 提取文件名信息
                extracted_info = extract_info_from_filename(original_name)

                # 匹配学生信息
                matched_student, match_reason = match_student_info(extracted_info, student_data, match_strategies)

                if not matched_student:
                    results.append({
                        "original": original_name,
                        "new": "未处理",
                        "status": f"未找到匹配的学生信息",
                        "extracted": str(extracted_info)
                    })
                    continue

                # 生成新文件名
                # 如果用户输入的 name_format 中已经包含了"实验（训）报告【电子版】-"，就不再添加
                if "实验（训）报告【电子版】-" in name_format or name_format.startswith("实验"):
                    # 用户自己已经编写了完整的后缀
                    if matched_student['id']:
                        new_filename = f"{matched_student['id']} {matched_student['name']} {name_format}{word_file.suffix}"
                    else:
                        new_filename = f"{matched_student['name']} {name_format}{word_file.suffix}"
                else:
                    # 用户只输入了简单的后缀（如"简历"），直接使用
                    if matched_student['id']:
                        new_filename = f"{matched_student['id']} {matched_student['name']} {name_format}{word_file.suffix}"
                    else:
                        new_filename = f"{matched_student['name']} {name_format}{word_file.suffix}"

                new_filename = clean_filename(new_filename)
                new_path = word_folder / new_filename

                # 处理重名
                if new_path.exists() and not new_path.samefile(word_file):
                    counter = 1
                    while new_path.exists():
                        name_without_ext = f"{matched_student['id']} {matched_student['name']} {name_format}({counter})" if \
                        matched_student[
                            'id'] else f"{matched_student['name']} {name_format}({counter})"
                        new_filename = clean_filename(name_without_ext) + word_file.suffix
                        new_path = word_folder / new_filename
                        counter += 1
                elif new_path.exists() and new_path.samefile(word_file):
                    results.append({
                        "original": original_name,
                        "new": "未处理",
                        "status": "文件名已经正确",
                        "match_reason": match_reason
                    })
                    continue

                # 执行重命名
                word_file.rename(new_path)
                success_count += 1

                results.append({
                    "original": original_name,
                    "new": new_filename,
                    "status": "成功",
                    "match_reason": match_reason
                })

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
        name_range = data.get('name_range')
        id_range = data.get('id_range')
        match_strategies = data.get('match_strategies', ['exact_name', 'partial_name', 'exact_id'])

        if not all([excel_path, word_folder, name_format, name_range]):
            return jsonify({"success": False, "message": "请填写必需字段"})

        if not os.path.exists(excel_path):
            return jsonify({"success": False, "message": f"Excel文件不存在: {excel_path}"})

        if not os.path.exists(word_folder):
            return jsonify({"success": False, "message": f"文件夹不存在: {word_folder}"})

        result = rename_word_files(excel_path, word_folder, name_format, name_range, id_range, match_strategies)
        return jsonify(result)

    except Exception as e:
        logger.error(f"API处理错误: {str(e)}")
        return jsonify({"success": False, "message": f"服务器错误: {str(e)}"})


# 简化的HTML模板
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>批量重命名工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
        h1 { text-align: center; margin-bottom: 30px; color: #333; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; color: #555; }
        input, select { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; }
        input:focus, select:focus { outline: none; border-color: #007bff; }
        .btn { background: #007bff; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; width: 100%; font-size: 16px; }
        .btn:hover { background: #0056b3; }
        .btn:disabled { background: #ccc; cursor: not-allowed; }
        .result { margin-top: 20px; padding: 15px; border-radius: 4px; display: none; }
        .result.success { background: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
        .result.error { background: #f8d7da; border: 1px solid #f5c6cb; color: #721c24; }
        .details { max-height: 300px; overflow-y: auto; margin-top: 10px; font-size: 13px; }
        .item { padding: 8px 0; border-bottom: 1px solid #eee; }
        .help { font-size: 12px; color: #666; margin-top: 3px; }
        .checkbox-group { display: flex; flex-wrap: wrap; gap: 15px; margin-top: 5px; }
        .checkbox-group input[type="checkbox"] { width: auto; margin-right: 5px; }
        .loading { display: none; text-align: center; margin-top: 15px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>批量文件重命名工具</h1>

        <form id="renameForm">
            <div class="form-group">
                <label>文件名后缀:</label>
                <input type="text" id="nameFormat" placeholder="例如：简历 或 实验（训）报告【电子版】-实验九" required>
                <div class="help">输入文件后缀，如"简历"或完整的"实验（训）报告【电子版】-实验九"</div>
            </div>

            <div class="form-group">
                <label>Excel文件路径:</label>
                <input type="text" id="excelPath" placeholder="C:\\path\\to\\file.xlsx" required>
            </div>

            <div class="form-group">
                <label>文件夹路径:</label>
                <input type="text" id="wordFolder" placeholder="C:\\path\\to\\folder" required>
                <div class="help">支持Word文档(.doc/.docx)、PDF文件(.pdf)和图片(.jpg/.png/.gif等)</div>
            </div>

            <div class="form-group">
                <label>姓名数据范围:</label>
                <input type="text" id="nameRange" placeholder="A2:A50" required>
                <div class="help">Excel中姓名所在的单元格范围，如 A2:A50</div>
            </div>

            <div class="form-group">
                <label>学号数据范围:</label>
                <input type="text" id="idRange" placeholder="C2:C50">
                <div class="help">Excel中学号所在的单元格范围，如 C2:C50（可选）</div>
            </div>

            <div class="form-group">
                <label>匹配策略:</label>
                <div class="checkbox-group">
                    <label><input type="checkbox" value="exact_name" checked> 精确姓名匹配</label>
                    <label><input type="checkbox" value="partial_name" checked> 部分姓名匹配</label>
                    <label><input type="checkbox" value="split_name"> 拆分姓名匹配</label>
                    <label><input type="checkbox" value="exact_id" checked> 精确学号匹配</label>
                    <label><input type="checkbox" value="partial_id"> 部分学号匹配</label>
                </div>
            </div>

            <button type="submit" class="btn" id="submitBtn">开始处理</button>
        </form>

        <div class="loading" id="loading">处理中，请稍候...</div>
        <div class="result" id="result">
            <div id="resultMessage"></div>
            <div class="details" id="resultDetails"></div>
        </div>
    </div>

    <script>
        document.getElementById('renameForm').addEventListener('submit', async function(e) {
            e.preventDefault();

            const submitBtn = document.getElementById('submitBtn');
            const loading = document.getElementById('loading');
            const result = document.getElementById('result');

            // 获取匹配策略
            const strategies = Array.from(document.querySelectorAll('input[type="checkbox"]:checked'))
                .map(cb => cb.value);

            const formData = {
                name_format: document.getElementById('nameFormat').value.trim(),
                excel_path: document.getElementById('excelPath').value.trim(),
                word_folder: document.getElementById('wordFolder').value.trim(),
                name_range: document.getElementById('nameRange').value.trim(),
                id_range: document.getElementById('idRange').value.trim(),
                match_strategies: strategies
            };

            if (!formData.name_format || !formData.excel_path || !formData.word_folder || !formData.name_range) {
                alert('请填写必需字段！');
                return;
            }

            if (strategies.length === 0) {
                alert('请至少选择一种匹配策略！');
                return;
            }

            submitBtn.disabled = true;
            submitBtn.textContent = '处理中...';
            loading.style.display = 'block';
            result.style.display = 'none';

            try {
                const response = await fetch('/rename', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(formData)
                });

                const data = await response.json();

                result.style.display = 'block';
                result.className = 'result ' + (data.success ? 'success' : 'error');

                if (data.success) {
                    document.getElementById('resultMessage').innerHTML = 
                        `<strong>${data.message}</strong><br>总文件数: ${data.total} | 成功: ${data.success_count}`;

                    if (data.results) {
                        let details = '';
                        data.results.forEach(item => {
                            details += `<div class="item">
                                <strong>原:</strong> ${item.original}<br>
                                <strong>新:</strong> ${item.new}<br>
                                <strong>状态:</strong> ${item.status}
                                ${item.match_reason ? `<br><strong>匹配:</strong> ${item.match_reason}` : ''}
                                ${item.extracted ? `<br><strong>提取:</strong> ${item.extracted}` : ''}
                            </div>`;
                        });
                        document.getElementById('resultDetails').innerHTML = details;
                    }
                } else {
                    document.getElementById('resultMessage').innerHTML = `<strong>错误:</strong> ${data.message}`;
                }

            } catch (error) {
                result.style.display = 'block';
                result.className = 'result error';
                document.getElementById('resultMessage').innerHTML = `<strong>网络错误:</strong> ${error.message}`;
            }

            submitBtn.disabled = false;
            submitBtn.textContent = '开始处理';
            loading.style.display = 'none';
        });
    </script>
</body>
</html>
'''

if __name__ == '__main__':
    print("工具已启动...")
    print("访问: http://localhost:5002")
    app.run(host='0.0.0.0', port=5002, debug=True)