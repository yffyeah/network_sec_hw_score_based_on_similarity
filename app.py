from flask import Flask, request, jsonify, send_from_directory
import os
import zipfile
import shutil
import subprocess
import sys
import re
import datetime

app = Flask(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'zip'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024  # 5MB文件大小限制

# 确保上传目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 计数器存储
COUNTERS_FILE = 'counters.json'

# 读取计数器
import json
def read_counters():
    if os.path.exists(COUNTERS_FILE):
        with open(COUNTERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'date': datetime.date.today().isoformat(), 'global_count': 0, 'ip_counts': {}}

# 保存计数器
def save_counters(counters):
    with open(COUNTERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(counters, f, ensure_ascii=False, indent=2)

# 检查并更新计数器
def check_and_update_counters():
    counters = read_counters()
    today = datetime.date.today().isoformat()
    
    # 如果日期不同，重置计数器
    if counters['date'] != today:
        counters = {'date': today, 'global_count': 0, 'ip_counts': {}}
    
    return counters

# 检查使用限制
def check_usage_limits():
    counters = check_and_update_counters()
    client_ip = request.remote_addr
    
    # 检查全局限制
    if counters['global_count'] >= 100:
        return False, '网站今日使用次数已达上限（100次）'
    
    # 检查IP限制
    if client_ip in counters['ip_counts'] and counters['ip_counts'][client_ip] >= 10:
        return False, '您今日使用次数已达上限（10次）'
    
    return True, None

# 更新计数器
def update_counters():
    counters = check_and_update_counters()
    client_ip = request.remote_addr
    
    # 更新全局计数
    counters['global_count'] += 1
    
    # 更新IP计数
    if client_ip not in counters['ip_counts']:
        counters['ip_counts'][client_ip] = 0
    counters['ip_counts'][client_ip] += 1
    
    save_counters(counters)
    return counters

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        # 检查使用限制
        allowed, message = check_usage_limits()
        if not allowed:
            return jsonify({'error': message})
        
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No selected file'})

        # 获取语言参数和评估方式
        language = request.form.get('language', 'chinese')
        evaluation = request.form.get('evaluation', 'text')
        print(f"language: {language}, evaluation: {evaluation}")
        
        if file and file.filename.endswith('.zip'):
            # 保存上传的文件
            filename = file.filename
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # 解压文件
            extract_dir = os.path.join(app.config['UPLOAD_FOLDER'], filename.replace('.zip', ''))
            if os.path.exists(extract_dir):
                shutil.rmtree(extract_dir)
            os.makedirs(extract_dir)
            
            # 处理ZIP文件编码问题，跳过隐藏文件
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                for zip_info in zip_ref.infolist():
                    # 跳过隐藏文件（以.开头）
                    basename = os.path.basename(zip_info.filename)
                    if basename.startswith('.'):
                        continue
                        
                    try:
                        # 读取文件内容
                        data = zip_ref.read(zip_info.filename)
                        # 获取文件名
                        original_name = zip_info.filename
                        # 尝试处理编码问题
                        try:
                            # 尝试用cp437解码（DOS/Win系统）
                            correct_name = original_name.encode('cp437').decode('utf-8', errors='ignore')
                        except:
                            correct_name = original_name
                        
                        # 如果文件名包含路径，只取最后一部分
                        if '/' in correct_name:
                            correct_name = correct_name.split('/')[-1]
                        
                        # 写入文件
                        outpath = os.path.join(extract_dir, correct_name)
                        with open(outpath, 'wb') as f:
                            f.write(data)
                    except Exception as e:
                        print(f"解压文件失败: {zip_info.filename}, 错误: {str(e)}")
            
            # 执行score_class.py，传递语言参数和评估方式
            try:
                result = subprocess.run(
                    ['python3', 'score_class.py', extract_dir, '--language', language, '--evaluation', evaluation],
                    capture_output=True,
                    text=True,
                    cwd=os.path.dirname(os.path.abspath(__file__))
                )
                # 添加以下两行
                print("score_class.py stdout:")
                print(result.stdout)
                print("score_class.py stderr:")
                print(result.stderr)
                sys.stdout.flush()

                # 查找生成的summary文件
                summary_files = []
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.startswith('summary_') and file.endswith('.txt'):
                            summary_files.append(os.path.join(root, file))

                # 查找回填成绩的Excel文件
                filled_excel_path = None
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        if file.endswith('_updated.xls'):
                            filled_excel_path = os.path.join(os.path.basename(extract_dir), file)
                            break
                    if filled_excel_path:
                        break

                # 读取summary文件内容
                summaries = {}
                for summary_file in summary_files:
                    folder_name = os.path.basename(os.path.dirname(summary_file))
                    # 尝试不同编码读取文件
                    encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']
                    content = None
                    for encoding in encodings:
                        try:
                            with open(summary_file, 'r', encoding=encoding) as f:
                                content = f.read()
                            break
                        except:
                            continue
                    
                    if content is None:
                        # 如果所有编码都失败，使用二进制模式读取
                        with open(summary_file, 'rb') as f:
                            content = f.read().decode('utf-8', errors='ignore')
                    
                    summaries[folder_name] = content

                # 提取成绩信息
                student_scores = []
                for folder_name, summary_content in summaries.items():
                    # 提取学号、姓名、相似度和分数
                    lines = summary_content.split('\n')
                    in_ranking = False
                    for line in lines:
                        line = line.strip()
                        if '学生平均相似度排名' in line:
                            in_ranking = True
                            continue
                        if in_ranking and line.startswith('=='):
                            break
                        if in_ranking and line and not line.startswith('排名') and '---' not in line:
                            parts = line.split()
                            if len(parts) >= 6:
                                try:
                                    rank = parts[0]
                                    student_id = parts[1]
                                    name = parts[2]
                                    class_name = parts[3]
                                    avg_sim = parts[4]
                                    score = parts[5]
                                    student_scores.append({
                                        'rank': rank,
                                        'id': student_id,
                                        'name': name,
                                        'className': class_name,
                                        'avgSimilarity': avg_sim,
                                        'score': score
                                    })
                                except:
                                    pass

                # 更新使用计数器
                counters = update_counters()
                
                return jsonify({
                    'success': True,
                    'summaries': summaries,
                    'studentScores': student_scores,
                    'filledExcelPath': filled_excel_path,
                    'output': result.stdout,
                    'error': result.stderr if result.returncode != 0 else None,
                    'usage': {
                        'globalCount': counters['global_count'],
                        'ipCount': counters['ip_counts'].get(request.remote_addr, 0),
                        'globalLimit': 100,
                        'ipLimit': 10
                    }
                })

            except Exception as e:
                return jsonify({'error': str(e)}), 500

    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Invalid file type'})

@app.route('/get_document', methods=['POST'])
def get_document():
    try:
        data = request.get_json()
        student_id = data.get('studentId')
        class_name = data.get('className')
        
        if not student_id:
            return jsonify({'error': '缺少学生ID'}), 400
        
        # 在uploads目录中查找包含该学号的文件
        uploads_dir = app.config['UPLOAD_FOLDER']
        found_file = None
        
        for root, dirs, files in os.walk(uploads_dir):
            for file in files:
                if student_id in file and (file.endswith('.doc') or file.endswith('.docx')):
                    found_file = os.path.join(root, file)
                    break
            if found_file:
                break
        
        if found_file:
            # 获取文件名并处理编码
            filename = os.path.basename(found_file)
            
            # 尝试URL编码文件名
            from urllib.parse import quote
            encoded_filename = quote(filename)
            
            # 返回文件，设置正确的Content-Disposition头
            response = send_from_directory(
                os.path.dirname(found_file),
                filename,
                as_attachment=True,
                download_name=filename
            )
            response.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{encoded_filename}"
            return response
        else:
            return jsonify({'error': '未找到该学生的作业文件'}), 404

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download_filled', methods=['GET'])
def download_filled():
    try:
        file_path = request.args.get('path')
        if not file_path:
            return jsonify({'error': '缺少文件路径'}), 400

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_path)

        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404

        filename = os.path.basename(file_path)
        from urllib.parse import quote
        encoded_filename = quote(filename)

        response = send_from_directory(
            os.path.dirname(file_path),
            filename,
            as_attachment=True,
            download_name=filename
        )
        response.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{encoded_filename}"
        return response
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=False)
