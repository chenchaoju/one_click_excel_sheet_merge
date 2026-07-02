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

CN_MONTH_MAP = {
    '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
    '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
    '十一': 11, '十二': 12,
}

FULL_DATE_PATTERNS = [
    (r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$', 'ymd'),
    (r'^(\d{4})(\d{2})(\d{2})$', 'ymd'),
    (r'^(\d{4})年(\d{1,2})月(\d{1,2})日$', 'ymd'),
]

MONTH_DATE_PATTERNS = [
    (r'^(\d{4})年(\d{1,2})月$', 'ym_num'),
    (r'^(\d{4})年([一二三四五六七八九十]{1,3})月$', 'ym_cn'),
    (r'^(\d{1,2})月$', 'm_num'),
    (r'^([一二三四五六七八九十]{1,3})月$', 'm_cn'),
    (r'^(\d{1,2})月份$', 'm_num'),
    (r'^([一二三四五六七八九十]{1,3})月份$', 'm_cn'),
]


def _cn_month_to_int(name):
    return CN_MONTH_MAP.get(name.strip())


def _last_day_of_month(y, m):
    if m == 12:
        next_y, next_m = y + 1, 1
    else:
        next_y, next_m = y, m + 1
    from datetime import date
    return (date(next_y, next_m, 1) - date(y, m, 1)).days


def parse_sheet_date(sheet_name):
    """尝试把表名解析成日期。成功返回 datetime，失败返回 None。

    支持的格式：
      - 完整日期：YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD / YYYYMMDD / YYYY年M月D日
      - 月份级（含年份）：YYYY年M月 / YYYY年五月（取该月最后一天）
      - 月份级（无年份，默认当前年）：5月 / 五月 / 5月份 / 五月份（取该月最后一天）
    """
    name = str(sheet_name).strip()

    for pattern, kind in FULL_DATE_PATTERNS:
        m = re.match(pattern, name)
        if not m:
            continue
        try:
            if kind == 'ymd':
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                return datetime(y, mo, d)
        except ValueError:
            return None

    for pattern, kind in MONTH_DATE_PATTERNS:
        m = re.match(pattern, name)
        if not m:
            continue
        try:
            if kind == 'ym_num':
                y, mo = int(m.group(1)), int(m.group(2))
            elif kind == 'ym_cn':
                y = int(m.group(1))
                mo = _cn_month_to_int(m.group(2))
                if mo is None:
                    return None
            elif kind == 'm_num':
                y = datetime.now().year
                mo = int(m.group(1))
            elif kind == 'm_cn':
                y = datetime.now().year
                mo = _cn_month_to_int(m.group(1))
                if mo is None:
                    return None
            else:
                continue

            if not (1 <= mo <= 12):
                return None
            d = _last_day_of_month(y, mo)
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


def _default_browse_root():
    """选择一个可访问的起始目录。"""
    candidates = [
        os.path.expanduser('~'),
        os.getcwd(),
        '/',
    ]
    seen = set()
    for p in candidates:
        if not p or p in seen:
            continue
        seen.add(p)
        if os.path.isdir(p) and os.access(p, os.R_OK | os.X_OK):
            return p
    return '/'


@app.route('/api/browse', methods=['POST'])
def api_browse():
    """列出某目录下的子目录，便于在前端逐级选择文件夹。"""
    data = request.get_json(silent=True) or {}
    path = (data.get('path') or '').strip()

    if not path:
        path = _default_browse_root()

    if not os.path.isdir(path):
        return jsonify({'error': f'路径不存在：{path}'}), 400

    if not os.access(path, os.R_OK):
        return jsonify({
            'error': f'无权限访问该目录：{path}',
            'current': path,
            'parent': os.path.dirname(path) if os.path.dirname(path) != path else None,
            'dirs': [],
            'permission_denied': True,
        }), 200

    try:
        entries = sorted(os.listdir(path))
    except PermissionError:
        return jsonify({
            'error': f'无权限读取目录内容：{path}',
            'current': path,
            'parent': os.path.dirname(path) if os.path.dirname(path) != path else None,
            'dirs': [],
            'permission_denied': True,
        }), 200

    dirs = []
    skipped = 0
    for name in entries:
        if name.startswith('.'):
            continue
        full = os.path.join(path, name)
        try:
            if os.path.isdir(full) and os.access(full, os.R_OK | os.X_OK):
                dirs.append({'name': name, 'path': full})
            else:
                skipped += 1
        except (PermissionError, OSError):
            skipped += 1
            continue

    parent = os.path.dirname(path)
    parent = parent if parent and parent != path else None

    result = {
        'current': path,
        'parent': parent,
        'dirs': dirs,
    }
    if skipped > 0:
        result['skipped'] = skipped
        result['hint'] = f'有 {skipped} 个项目因权限不足或不可读而被跳过'
    return jsonify(result)


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
