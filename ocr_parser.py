# -*- coding: utf-8 -*-
"""
OCR表格解析模块
从图片中识别表格内容，提取指定人员的工作内容
"""
import re
import os
from rapidocr_onnxruntime import RapidOCR
from PIL import Image

# 全局OCR引擎（只初始化一次）
_global_ocr = None


def get_global_ocr():
    """获取全局OCR引擎，懒加载"""
    global _global_ocr
    if _global_ocr is None:
        _global_ocr = RapidOCR()
    return _global_ocr


def preprocess_image(image_path, max_width=1200):
    """图片预处理：缩小尺寸，加快OCR速度。返回处理后的图片路径"""
    try:
        img = Image.open(image_path)
        w, h = img.size
        if w > max_width:
            ratio = max_width / w
            new_h = int(h * ratio)
            img = img.resize((max_width, new_h), Image.LANCZOS)
            # 保存到临时文件
            base, ext = os.path.splitext(image_path)
            new_path = f"{base}_small{ext}"
            img.save(new_path, quality=95)
            return new_path
        return image_path
    except Exception:
        return image_path


class OCRParser:
    def __init__(self, target_name="陈子怡", text_corrections=None, use_global_ocr=True):
        """
        Args:
            target_name: 要提取的人员姓名
            text_corrections: OCR常见错误替换字典 {错误词: 正确词}
            use_global_ocr: 是否使用全局共享的OCR引擎
        """
        self.target_name = target_name
        self.text_corrections = text_corrections or {}
        self.use_global_ocr = use_global_ocr
        self._local_ocr = None

    @property
    def ocr(self):
        if self.use_global_ocr:
            return get_global_ocr()
        if self._local_ocr is None:
            self._local_ocr = RapidOCR()
        return self._local_ocr

    def _correct_text(self, text):
        """应用OCR常见错误修正"""
        for wrong, right in self.text_corrections.items():
            text = text.replace(wrong, right)
        return text

    def _get_text_blocks(self, image_path):
        """OCR识别，返回文本块列表 [(center_x, center_y, text), ...]"""
        result, _ = self.ocr(image_path)
        blocks = []
        if result is None:
            return blocks
        for item in result:
            box, text, score = item
            text = self._correct_text(text.strip())
            if not text:
                continue
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            cx = sum(xs) / 4
            cy = sum(ys) / 4
            # 计算块的高度，用于行聚类阈值
            height = max(ys) - min(ys)
            blocks.append({
                'cx': cx,
                'cy': cy,
                'text': text,
                'height': height,
                'x_min': min(xs),
                'x_max': max(xs),
                'y_min': min(ys),
                'y_max': max(ys),
            })
        return blocks

    def _detect_columns(self, blocks):
        """
        检测表格列边界。
        通过表头行（包含"日期""人员""工作内容""完成时间"等关键词）确定列的x范围。
        返回: {'date': (x_min, x_max), 'person': ..., 'content': ..., 'deadline': ...}
        """
        # 找到表头行：y最小的一批文本中包含关键词
        header_keywords = ['日期', '人员', '工作内容', '完成时间']
        header_blocks = []
        for b in blocks:
            for kw in header_keywords:
                if kw in b['text']:
                    header_blocks.append(b)
                    break

        if not header_blocks:
            #  fallback：用x坐标聚类
            return self._detect_columns_by_clustering(blocks)

        # 按x排序表头，确定列顺序
        header_blocks.sort(key=lambda b: b['cx'])
        col_order = []
        header_map = {}
        for hb in header_blocks:
            if '日期' in hb['text']:
                col_order.append('date')
                header_map['date'] = hb
            elif '人员' in hb['text']:
                col_order.append('person')
                header_map['person'] = hb
            elif '工作内容' in hb['text']:
                col_order.append('content')
                header_map['content'] = hb
            elif '完成时间' in hb['text']:
                col_order.append('deadline')
                header_map['deadline'] = hb

        # 精确确定人员列范围：
        # 人名通常在人员列表头正下方较窄的x范围内
        person_hb = header_map.get('person')
        person_col_left = 0
        person_col_right = 0
        if person_hb:
            px = person_hb['cx']
            name_blocks = [b for b in blocks
                           if abs(b['cx'] - px) < 70
                           and b['text'] not in ('人员',)
                           and len(b['text']) <= 5]
            if name_blocks:
                person_col_left = min(b['x_min'] for b in name_blocks) - 20
                person_col_right = max(b['x_max'] for b in name_blocks) + 25
            else:
                person_col_left = px - 50
                person_col_right = px + 80

        # 精确确定完成时间列左边界
        deadline_hb = header_map.get('deadline')
        deadline_col_left = 99999
        if deadline_hb:
            dx = deadline_hb['cx']
            deadline_text_blocks = [b for b in blocks
                                    if abs(b['cx'] - dx) < 100
                                    and b['text'] not in ('完成时间节点', '完成时间')]
            if deadline_text_blocks:
                deadline_col_left = min(b['x_min'] for b in deadline_text_blocks) - 15
            else:
                deadline_col_left = dx - 60

        # 最终列边界
        columns = {}
        for i, name in enumerate(col_order):
            if name == 'date':
                x_min = 0
                x_max = person_col_left  # 日期列到人员列左边界为止
            elif name == 'person':
                x_min = person_col_left
                x_max = person_col_right
            elif name == 'content':
                x_min = person_col_right
                x_max = deadline_col_left
            elif name == 'deadline':
                x_min = deadline_col_left
                x_max = 99999
            else:
                x_min, x_max = 0, 99999
            columns[name] = (x_min, x_max)

        return columns

    def _detect_columns_by_clustering(self, blocks):
        """无表头时的fallback：按x坐标聚类"""
        xs = sorted([b['cx'] for b in blocks])
        # 简单分为4列
        n = len(xs)
        q1 = xs[n // 4] if n > 4 else xs[0]
        q2 = xs[n // 2] if n > 4 else xs[-1]
        q3 = xs[3 * n // 4] if n > 4 else xs[-1]
        return {
            'date': (0, (q1 + q2) / 2),
            'person': ((q1 + q2) / 2, (q2 + q3) / 2),
            'content': ((q2 + q3) / 2, (q3 + xs[-1]) / 2),
            'deadline': ((q3 + xs[-1]) / 2, 99999),
        }

    def _cluster_rows(self, blocks):
        """
        按y坐标聚类行。
        返回: [[block, ...], ...] 每行一个列表
        """
        if not blocks:
            return []

        # 按y排序
        sorted_blocks = sorted(blocks, key=lambda b: b['cy'])

        # 估算平均行高
        heights = [b['height'] for b in sorted_blocks if b['height'] > 0]
        avg_height = sum(heights) / len(heights) if heights else 20
        row_threshold = avg_height * 0.8  # 行间距阈值

        rows = []
        current_row = [sorted_blocks[0]]
        current_y = sorted_blocks[0]['cy']

        for b in sorted_blocks[1:]:
            if abs(b['cy'] - current_y) < row_threshold:
                current_row.append(b)
                current_y = sum(x['cy'] for x in current_row) / len(current_row)
            else:
                rows.append(current_row)
                current_row = [b]
                current_y = b['cy']
        rows.append(current_row)

        return rows

    def _assign_to_columns(self, row, columns):
        """将一行中的文本块分配到各列"""
        result = {name: [] for name in columns}
        for b in row:
            for name, (x_min, x_max) in columns.items():
                if x_min <= b['cx'] < x_max:
                    result[name].append(b)
                    break
        # 每列内按x排序
        for name in result:
            result[name].sort(key=lambda b: b['cx'])
        return result

    def _extract_date(self, date_blocks):
        """从日期列文本块中提取日期字符串"""
        if not date_blocks:
            return None
        texts = [b['text'] for b in date_blocks]
        full = ' '.join(texts)
        # 匹配 2026/09/07 或 2026-09-07 或 9月7日 等格式
        m = re.search(r'(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})', full)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        m = re.search(r'(\d{1,2})月(\d{1,2})日', full)
        if m:
            return f"{m.group(1)}月{m.group(2)}日"
        return full.strip()

    def _extract_weekday(self, date_blocks):
        """提取星期信息"""
        if not date_blocks:
            return None
        full = ' '.join(b['text'] for b in date_blocks)
        m = re.search(r'周([一二三四五六日天])', full)
        if m:
            return m.group(1)
        m = re.search(r'（周([一二三四五六日天])）', full)
        if m:
            return m.group(1)
        return None

    def _is_valid_content(self, text):
        """判断是否为有效的工作内容文本（过滤OCR噪声）"""
        if not text or len(text) < 2:
            return False
        # 有效字符：中文、字母、数字
        valid = sum(1 for c in text if '\u4e00' <= c <= '\u9fff' or c.isalnum())
        ratio = valid / len(text)
        # 有效字符比例低于40%视为噪声
        if ratio < 0.4:
            return False
        # 纯符号或纯括号的过滤
        if re.match(r'^[\(\)\[\]\/\\|，。、\s]+$', text):
            return False
        return True

    def _merge_content_blocks(self, content_blocks):
        """
        合并工作内容列的文本块。
        工作内容列可能包含编号(1、2、)和具体内容，需要按行内的y坐标进一步分组，
        然后合并为完整的工作内容条目。
        """
        if not content_blocks:
            return []

        # 先过滤掉OCR噪声
        content_blocks = [b for b in content_blocks if self._is_valid_content(b['text'])]
        if not content_blocks:
            return []

        # 工作内容列可能有多行（每人有多个工作项），需要按y子聚类
        # 先按y排序
        sorted_blocks = sorted(content_blocks, key=lambda b: b['cy'])

        # 估算子行高
        heights = [b['height'] for b in sorted_blocks if b['height'] > 0]
        avg_h = sum(heights) / len(heights) if heights else 18
        sub_threshold = avg_h * 0.7

        # 子行聚类
        sub_rows = []
        current = [sorted_blocks[0]]
        current_y = sorted_blocks[0]['cy']
        for b in sorted_blocks[1:]:
            if abs(b['cy'] - current_y) < sub_threshold:
                current.append(b)
                current_y = sum(x['cy'] for x in current) / len(current)
            else:
                sub_rows.append(current)
                current = [b]
                current_y = b['cy']
        sub_rows.append(current)

        # 每个子行合并为一条文本（编号+内容）
        items = []
        for sr in sub_rows:
            sr.sort(key=lambda b: b['cx'])
            # 合并文本，编号和内容之间加空格
            texts = [b['text'] for b in sr]
            # 如果第一个文本看起来是编号（如"1、" "2、" "1>" "1HCT"），和后面的合并
            merged = ''.join(texts)
            # 清理：编号后缺少顿号的补上
            merged = re.sub(r'^(\d+)(?=[A-Z\u4e00-\u9fff])', r'\1、', merged)
            merged = re.sub(r'^(\d+)>', r'\1、', merged)
            items.append(merged)

        return items

    def parse_image(self, image_path):
        """
        解析单张图片，提取目标人员的工作内容。
        采用按人员y范围收集的策略，不依赖全局行聚类。

        Returns:
            dict: {
                'date': '2026/09/07',
                'weekday': '一',
                'person': '陈子怡',
                'work_items': ['1、xxx', '2、xxx'],
                'deadline': '9月7日',
                'found': True/False
            }
        """
        # 图片预处理：缩小尺寸加快OCR
        processed_path = preprocess_image(image_path)
        blocks = self._get_text_blocks(processed_path)
        # 清理临时文件
        if processed_path != image_path and os.path.exists(processed_path):
            try:
                os.remove(processed_path)
            except Exception:
                pass
        if not blocks:
            return {'found': False, 'reason': 'OCR未识别到文本', 'image': image_path}

        columns = self._detect_columns(blocks)

        # 将所有块按列分类
        col_blocks = {name: [] for name in columns}
        for b in blocks:
            for name, (x_min, x_max) in columns.items():
                if x_min <= b['cx'] < x_max:
                    col_blocks[name].append(b)
                    break

        # 人员列：找到所有人员及其y坐标（排除表头）
        person_blocks = sorted(
            [b for b in col_blocks.get('person', []) if b['text'] not in ('人员',)],
            key=lambda b: b['cy']
        )
        if not person_blocks:
            return {'found': False, 'reason': '未识别到人员列', 'image': image_path}

        # 找到目标人员
        target_blocks = [b for b in person_blocks if self.target_name in b['text']]
        if not target_blocks:
            return {'found': False, 'reason': f'未找到{self.target_name}', 'image': image_path}

        target_block = target_blocks[0]
        target_y = target_block['cy']

        # 确定目标人员的y范围：上一个人员中心 到 下一个人员中心
        person_ys = [b['cy'] for b in person_blocks]
        idx = person_ys.index(target_y) if target_y in person_ys else 0

        # 估算平均行高（相邻人员间距）
        if len(person_ys) >= 2:
            gaps = [person_ys[i+1] - person_ys[i] for i in range(len(person_ys)-1)]
            avg_gap = sum(gaps) / len(gaps)
        else:
            avg_gap = 40

        if idx > 0:
            y_min = (person_ys[idx - 1] + target_y) / 2
        else:
            y_min = 0

        if idx < len(person_ys) - 1:
            y_max = (target_y + person_ys[idx + 1]) / 2
        else:
            # 最后一人：限制y范围，避免包含表格外噪声
            y_max = target_y + avg_gap * 1.5

        # 收集该y范围内的工作内容块
        content_blocks = [b for b in col_blocks.get('content', [])
                          if y_min <= b['cy'] <= y_max]
        content_blocks.sort(key=lambda b: (b['cy'], b['cx']))

        work_items = self._merge_content_blocks(content_blocks)

        # 收集完成时间
        deadline_blocks = [b for b in col_blocks.get('deadline', [])
                           if y_min <= b['cy'] <= y_max]
        deadline = ' '.join(b['text'] for b in deadline_blocks).strip() if deadline_blocks else ''

        # 日期列：日期是合并单元格，一张图片只有一个日期。
        # 收集所有日期列块（含y范围内和上方的），从中提取日期和星期
        date_blocks = list(col_blocks.get('date', []))
        # 排除表头
        date_blocks = [b for b in date_blocks if b['text'] not in ('日期',)]

        date_str = self._extract_date(date_blocks)
        weekday = self._extract_weekday(date_blocks)

        # 如果没提取到日期，尝试从完成时间推断（如"9月7日"）
        if not date_str and deadline:
            m = re.search(r'(\d{1,2})月(\d{1,2})日', deadline)
            if m:
                date_str = f"{m.group(1)}月{m.group(2)}日"

        # 如果还是没日期，从同列其他行的完成时间推断（取出现最多的日期）
        if not date_str:
            all_deadlines = [b['text'] for b in col_blocks.get('deadline', [])]
            date_counts = {}
            for dt in all_deadlines:
                m = re.search(r'(\d{1,2})月(\d{1,2})日', dt)
                if m:
                    key = f"{m.group(1)}月{m.group(2)}日"
                    date_counts[key] = date_counts.get(key, 0) + 1
            if date_counts:
                date_str = max(date_counts, key=date_counts.get)

        return {
            'found': True,
            'date': date_str,
            'weekday': weekday,
            'person': self.target_name,
            'work_items': work_items,
            'deadline': deadline,
            'image': image_path,
        }

    def parse_images(self, image_paths):
        """批量解析多张图片，按日期排序"""
        results = []
        for path in image_paths:
            r = self.parse_image(path)
            results.append(r)
        # 按日期排序
        def date_key(r):
            d = r.get('date')
            if not d:
                # 从完成时间推断
                deadline = r.get('deadline') or ''
                m = re.search(r'(\d{1,2})月(\d{1,2})日', str(deadline))
                if m:
                    return (2026, int(m.group(1)), int(m.group(2)))
                return (9999, 99, 99)
            if not isinstance(d, str):
                d = str(d)
            m = re.search(r'(\d{4})/(\d{1,2})/(\d{1,2})', d)
            if m:
                return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
            m = re.search(r'(\d{1,2})月(\d{1,2})日', d)
            if m:
                return (2026, int(m.group(1)), int(m.group(2)))
            return (9999, 99, 99)
        results.sort(key=date_key)
        return results
