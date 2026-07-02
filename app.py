# -*- coding: utf-8 -*-
"""合并 Excel 文件夹中所有 Excel 表里最后一个日期子表的 Flask 应用。"""
import io
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)

# 支持的日期命名格式：
#   YYYY-MM-DD  /  YYYY/MM/DD  /  YYYY.MM.DD
#   YYYYMMDD
#   YYYY年MM月DD日 / YYYY年M月D日
DATE_PATTERNS = [
    r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$',
    r'^(\d{4})(\d{2})(\d{2})$',
    r'^(\d{4})年(\d{1,2})月(\d{1,2})日$',
]


def parse_sheet_date(sheet_name):
    """尝试把表名解析成日期。成功返回 datetime，失败返回 None。"""
    name = str(sheet_name).strip()
    for pattern in DATE_PATTERNS:
        m = re.match(pattern, name)
        if m:
            try:
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                return datetime(y, mo, d)
            except ValueError:
                return None
    return None


def find_latest_date_sheet(file_path):
    """找出 Excel 文件中日期最大的子表。

    返回 (latest_tuple, error_message)。
    latest_tuple = (datetime, sheet_name)；找不到时为 None。
    """
    try:
        xl = pd.ExcelFile(file_path)
    except Exception as e:  # noqa: BLE001
        return None, f'无法读取文件: {e}'

    dated_sheets = []
    for sheet in xl.sheet_names:
        dt = parse_sheet_date(sheet)
        if dt is not None:
            dated_sheets.append((dt, sheet))

    if not dated_sheets:
        return None, '没有找到日期命名的子表'

    dated_sheets.sort(key=lambda x: x[0])
    return dated_sheets[-1], None


def scan_folder(folder):
    """扫描文件夹下所有 Excel 文件，返回每个文件的预览信息。"""
    excel_files = []
    for ext in ('*.xlsx', '*.xls'):
        excel_files.extend(Path(folder).glob(ext))

    items = []
    for fp in sorted(excel_files):
        latest, err = find_latest_date_sheet(fp)
        if latest is None:
            items.append({
                'file': fp.name,
                'path': str(fp),
                'status': 'skipped',
                'reason': err,
            })
        else:
            dt, sheet = latest
            items.append({
                'file': fp.name,
                'path': str(fp),
                'status': 'ok',
                'sheet': sheet,
                'date': dt.strftime('%Y-%m-%d'),
            })
    return items


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/scan', methods=['POST'])
def api_scan():
    data = request.get_json(silent=True) or {}
    folder = (data.get('folder') or '').strip()
    if not folder or not os.path.isdir(folder):
        return jsonify({'error': '无效的文件夹路径'}), 400

    items = scan_folder(folder)
    ok_count = sum(1 for i in items if i['status'] == 'ok')
    return jsonify({
        'folder': folder,
        'total': len(items),
        'ok': ok_count,
        'skipped': len(items) - ok_count,
        'items': items,
    })


@app.route('/api/browse', methods=['POST'])
def api_browse():
    """列出某目录下的子目录，便于在前端逐级选择文件夹。"""
    data = request.get_json(silent=True) or {}
    path = (data.get('path') or '').strip()

    if not path:
        # 默认从用户主目录开始
        path = os.path.expanduser('~')
    if not os.path.isdir(path):
        return jsonify({'error': '路径不存在'}), 400

    try:
        entries = sorted(os.listdir(path))
    except PermissionError:
        return jsonify({'error': '无权限访问该目录'}), 403

    dirs = []
    for name in entries:
        full = os.path.join(path, name)
        if os.path.isdir(full) and not name.startswith('.'):
            dirs.append({'name': name, 'path': full})

    parent = os.path.dirname(path)
    return jsonify({
        'current': path,
        'parent': parent if parent and parent != path else None,
        'dirs': dirs,
    })


@app.route('/api/merge', methods=['POST'])
def api_merge():
    data = request.get_json(silent=True) or {}
    folder = (data.get('folder') or '').strip()
    if not folder or not os.path.isdir(folder):
        return jsonify({'error': '无效的文件夹路径'}), 400

    items = scan_folder(folder)
    frames = []
    report = []
    for it in items:
        if it['status'] != 'ok':
            report.append(it)
            continue
        try:
            df = pd.read_excel(it['path'], sheet_name=it['sheet'])
        except Exception as e:  # noqa: BLE001
            it['status'] = 'error'
            it['reason'] = f'读取数据失败: {e}'
            report.append(it)
            continue

        if df.empty:
            it['status'] = 'skipped'
            it['reason'] = '子表为空'
            report.append(it)
            continue

        df.insert(0, '_源文件', it['file'])
        df.insert(1, '_子表名', it['sheet'])
        df.insert(2, '_日期', it['date'])
        frames.append(df)
        it['rows'] = len(df)
        report.append(it)

    if not frames:
        return jsonify({'error': '没有可合并的数据', 'details': report}), 400

    merged = pd.concat(frames, ignore_index=True, sort=False)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        merged.to_excel(writer, index=False, sheet_name='合并结果')
    output.seek(0)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'合并结果_{timestamp}.xlsx'

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
