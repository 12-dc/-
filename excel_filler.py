# -*- coding: utf-8 -*-
"""
Excel填入模块
将解析出的工作内容填入周报模板，保持原格式不变
"""
import re
import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


class ExcelFiller:
    def __init__(self, template_path, output_path, team_info=None):
        """
        Args:
            template_path: 周报模板路径
            output_path: 输出文件路径
            team_info: 配合部门及人员默认值
        """
        self.template_path = template_path
        self.output_path = output_path
        self.team_info = team_info or '合肥-残留事业部：汪芳、张翠、沈欣怡、陈晓雯、江梦蝶'
        self.wb = None
        self.ws = None
        self._date_col = 'A'
        self._weekday_col = 'B'
        self._category_col = 'C'
        self._contact_col = 'D'
        self._team_col = 'E'
        self._content_col = 'F'
        self._done_col = 'G'
        self._issue_col = 'H'
        self._data_start_row = None
        self._day_ranges = []  # [(start_row, end_row), ...] 每天的行范围
        self._summary_row = None
        self._next_week_start_row = None

    def _analyze_template(self):
        """分析模板结构，找到数据行范围和每天的行分组"""
        # 找到列标题行（包含"日期"的行）
        header_row = None
        for row in range(1, 10):
            cell_val = self.ws.cell(row=row, column=1).value
            if cell_val and '日期' in str(cell_val):
                header_row = row
                break
        if header_row is None:
            header_row = 3  # 默认第3行

        self._data_start_row = header_row + 1

        # 找到每天的行范围（通过A列合并单元格或日期值）
        # 先收集A列有值的行（合并单元格的左上角）
        date_rows = []
        for row in range(self._data_start_row, self.ws.max_row + 1):
            val = self.ws.cell(row=row, column=1).value
            if val is not None:
                s = str(val)
                # 先排除非日期的区域标题
                if any(kw in s for kw in ['小结', '帮助', '建议', '下周', '计划']):
                    break
                # 判断是否是日期（数字或日期字符串）
                if re.match(r'^\d+$', s) or re.match(r'\d{4}', s) or '月' in s:
                    date_rows.append(row)
                # 注意：不含'周'判断，避免"本周工作小结"被误判

        # 确定每天的行范围
        self._day_ranges = []
        for i, dr in enumerate(date_rows):
            end = date_rows[i + 1] - 1 if i + 1 < len(date_rows) else self._find_section_end(dr)
            self._day_ranges.append((dr, end))

        # 找到本周小结行
        for row in range(self._data_start_row, self.ws.max_row + 1):
            val = self.ws.cell(row=row, column=1).value
            if val and '小结' in str(val):
                self._summary_row = row
                break

        # 找到下周工作计划区域
        for row in range(self._data_start_row, self.ws.max_row + 1):
            val = self.ws.cell(row=row, column=1).value
            if val and '下周' in str(val):
                self._next_week_start_row = row
                break

    def _find_section_end(self, start_row):
        """找到某个区域的结束行（遇到下一个非空A列或小结区域）"""
        for row in range(start_row + 1, self.ws.max_row + 1):
            val = self.ws.cell(row=row, column=1).value
            if val is not None:
                s = str(val)
                if '小结' in s or '帮助' in s or '建议' in s or '下周' in s or re.match(r'^\d+$', s):
                    return row - 1
        return self.ws.max_row

    def _date_to_serial(self, date_str):
        """将日期字符串转为Excel序列号"""
        from datetime import datetime
        if not date_str:
            return None
        m = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})', date_str)
        if m:
            dt = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            # Excel日期序列号（1900系统，有1900闰年bug）
            delta = dt - datetime(1899, 12, 30)
            return delta.days
        m = re.match(r'(\d{1,2})月(\d{1,2})日', date_str)
        if m:
            dt = datetime(2026, int(m.group(1)), int(m.group(2)))
            delta = dt - datetime(1899, 12, 30)
            return delta.days
        return None

    def _weekday_to_char(self, weekday):
        """将星期转为单字"""
        if not weekday:
            return None
        mapping = {'一': '一', '二': '二', '三': '三', '四': '四', '五': '五', '六': '六', '日': '日', '天': '日'}
        for k, v in mapping.items():
            if k in weekday:
                return v
        return weekday

    def _categorize_work(self, work_items):
        """
        将工作项分类，返回 [(category, [items]), ...]
        分类规则：包含"样品检测"->样品检测；包含"预实验"或"空白制备"或"空白基质"->预实验；
        包含"面谈"->面谈；包含"前处理"->前处理；其他->其他
        """
        categories = []
        for item in work_items:
            cat = '其他'
            if '样品检测' in item or '检测进样' in item or '上样品检测' in item:
                cat = '样品检测'
            elif '预实验' in item or '空白制备' in item or '空白基质' in item or '基质制备' in item:
                cat = '预实验'
            elif '面谈' in item:
                cat = '面谈'
            elif '前处理' in item or '添加回收' in item:
                cat = '前处理'
            elif '请假' in item or '调休' in item:
                cat = '请假'
            elif '流转表' in item or '填写' in item:
                cat = '行政'

            # 合并同类
            found = False
            for c, items in categories:
                if c == cat:
                    items.append(item)
                    found = True
                    break
            if not found:
                categories.append((cat, [item]))

        return categories

    def _renumber_items(self, items):
        """对工作项重新编号，去掉原始编号，改为连续的1、2、3..."""
        renumbered = []
        for i, item in enumerate(items):
            # 去掉开头的编号（如"1、" "2、" "1>"等）
            cleaned = re.sub(r'^\d+[、>\.．]\s*', '', item)
            renumbered.append(f'{i+1}、{cleaned}')
        return renumbered

    def _fill_day(self, day_range, date_str, weekday, work_items):
        """填入一天的工作内容"""
        start_row, end_row = day_range
        available_rows = end_row - start_row + 1

        # 填入日期和星期
        if isinstance(date_str, (int, float)):
            # 已经是日期序列号
            self.ws.cell(row=start_row, column=1).value = int(date_str)
        else:
            date_serial = self._date_to_serial(date_str)
            if date_serial is not None:
                self.ws.cell(row=start_row, column=1).value = date_serial
            elif date_str:
                self.ws.cell(row=start_row, column=1).value = date_str

        wd = self._weekday_to_char(weekday)
        if not wd:
            # 从日期推断星期
            if isinstance(date_str, (int, float)):
                ds = int(date_str)
            else:
                ds = self._date_to_serial(date_str)
            if ds:
                # Excel序列号：1=1900-01-01(周日), 序列号mod 7: 0=周六,1=周日,2=周一...6=周五
                weekday_map = ['日', '一', '二', '三', '四', '五', '六']
                wd = weekday_map[(ds - 1) % 7]
        if wd:
            try:
                self.ws.cell(row=start_row, column=2).value = wd
            except AttributeError:
                pass  # B列是合并单元格的非左上角，跳过

        if not work_items:
            # 没有资料，清空内容行
            for r in range(start_row, end_row + 1):
                for col in [3, 4, 5, 6, 7, 8]:
                    # 只清空值，保留格式
                    self.ws.cell(row=r, column=col).value = None
            return

        # 分类
        categories = self._categorize_work(work_items)

        # 如果分类数 <= 可用行数，每类一行
        # 否则合并到较少的行中
        if len(categories) <= available_rows:
            # 每类一行，多余行留空
            for i, (cat, items) in enumerate(categories):
                row = start_row + i
                self.ws.cell(row=row, column=3).value = f'{i+1}、{cat}'
                self.ws.cell(row=row, column=4).value = '/'
                # 面谈类的配合人员特殊处理
                if cat == '面谈':
                    # 从工作项中提取面谈对象
                    team_val = '/'
                    for item in items:
                        m = re.search(r'面谈[辅导]*[-—](.+)', item)
                        if m:
                            team_val = m.group(1).strip()
                    self.ws.cell(row=row, column=5).value = team_val
                else:
                    self.ws.cell(row=row, column=5).value = self.team_info
                # 重新编号
                renumbered = self._renumber_items(items)
                self.ws.cell(row=row, column=6).value = '\n'.join(renumbered)
                self.ws.cell(row=row, column=7).value = '是'
                self.ws.cell(row=row, column=8).value = '/'
            # 清空多余行
            for r in range(start_row + len(categories), end_row + 1):
                for col in [3, 4, 5, 6, 7, 8]:
                    self.ws.cell(row=r, column=col).value = None
        else:
            # 分类多于可用行，合并所有内容到第一行，其余行留空
            all_items = []
            for cat, items in categories:
                all_items.extend(items)
            # 重新编号
            renumbered = self._renumber_items(all_items)
            self.ws.cell(row=start_row, column=3).value = '1、综合'
            self.ws.cell(row=start_row, column=4).value = '/'
            self.ws.cell(row=start_row, column=5).value = self.team_info
            self.ws.cell(row=start_row, column=6).value = '\n'.join(renumbered)
            self.ws.cell(row=start_row, column=7).value = '是'
            self.ws.cell(row=start_row, column=8).value = '/'
            for r in range(start_row + 1, end_row + 1):
                for col in [3, 4, 5, 6, 7, 8]:
                    self.ws.cell(row=r, column=col).value = None

    def _generate_summary(self, all_work_items):
        """生成本周工作小结"""
        # 统计项目数、药物数、基质数
        project_codes = set()
        all_drugs = set()
        all_matrices = set()
        category_count = {'样品检测': 0, '预实验': 0, '面谈': 0, '前处理': 0, '其他': 0}

        for date, items in all_work_items:
            for item in items:
                # 提取项目编号
                m = re.search(r'(HCT\d+[A-Z]*\d*)', item)
                if m:
                    project_codes.add(m.group(1))
                # 分类计数
                if '样品检测' in item or '检测进样' in item:
                    category_count['样品检测'] += 1
                elif '预实验' in item or '空白制备' in item or '空白基质' in item:
                    category_count['预实验'] += 1
                elif '面谈' in item:
                    category_count['面谈'] += 1
                elif '前处理' in item or '添加回收' in item:
                    category_count['前处理'] += 1
                else:
                    category_count['其他'] += 1

        lines = []
        idx = 1
        if category_count['样品检测'] > 0:
            lines.append(f'{idx}、{category_count["样品检测"]}项样品检测')
            idx += 1
        if category_count['预实验'] > 0:
            lines.append(f'{idx}、{category_count["预实验"]}项预实验及空白制备')
            idx += 1
        if category_count['前处理'] > 0:
            lines.append(f'{idx}、{category_count["前处理"]}项前处理')
            idx += 1
        if category_count['面谈'] > 0:
            lines.append(f'{idx}、{category_count["面谈"]}个面谈辅导')
            idx += 1
        if category_count['其他'] > 0:
            lines.append(f'{idx}、{category_count["其他"]}项其他工作')
            idx += 1

        return '\n'.join(lines) if lines else '/'

    def fill(self, parsed_results):
        """
        填入解析结果

        Args:
            parsed_results: [{'date':..., 'weekday':..., 'work_items':[...]}, ...]
        """
        self.wb = load_workbook(self.template_path)
        # 默认用第一个sheet
        self.ws = self.wb[self.wb.sheetnames[0]]

        self._analyze_template()

        # 按天填入
        all_work = []
        last_date_serial = None
        weekday_map = ['一', '二', '三', '四', '五', '六', '日']
        for i, day_range in enumerate(self._day_ranges):
            if i < len(parsed_results):
                r = parsed_results[i]
                if r.get('found'):
                    date_str = r.get('date')
                    date_serial = self._date_to_serial(date_str)
                    if date_serial:
                        last_date_serial = date_serial
                    self._fill_day(day_range, date_str, r.get('weekday'), r.get('work_items', []))
                    all_work.append((date_str, r.get('work_items', [])))
                else:
                    self._fill_day(day_range, r.get('date'), r.get('weekday'), [])
            else:
                # 模板有更多天但没有数据，根据前一天推断日期，内容留空
                inferred_date = None
                inferred_weekday = None
                if last_date_serial is not None:
                    inferred_date = last_date_serial + 1
                    last_date_serial = inferred_date
                    # 推断星期（Excel中1900-01-01是周日，序列号1对应周日）
                    # 序列号 mod 7：0=周六, 1=周日, 2=周一, ..., 6=周五
                    wd_idx = (inferred_date - 1) % 7
                    weekday_map_full = ['日', '一', '二', '三', '四', '五', '六']
                    inferred_weekday = weekday_map_full[wd_idx]
                self._fill_day(day_range, inferred_date, inferred_weekday, [])

        # 如果解析结果比模板天数多，需要插入行（暂不支持，记录警告）
        if len(parsed_results) > len(self._day_ranges):
            print(f'警告：解析到{len(parsed_results)}天数据，但模板只有{len(self._day_ranges)}天，超出部分未填入')

        # 更新本周小结
        if self._summary_row:
            summary = self._generate_summary(all_work)
            # 小结在C列（可能是合并单元格）
            self.ws.cell(row=self._summary_row, column=3).value = summary

        # 清空下周工作计划（无资料）
        if self._next_week_start_row:
            for row in range(self._next_week_start_row + 1, self.ws.max_row + 1):
                for col in range(1, 9):
                    cell = self.ws.cell(row=row, column=col)
                    # 保留表头行（计划执行时间等），清空数据行
                    val = cell.value
                    if val and ('计划执行时间' in str(val) or '计划工作项目' in str(val) or
                                '外部联络方式' in str(val) or '配合部门' in str(val) or
                                '计划简要描述' in str(val) or '可能存在问题' in str(val)):
                        continue
                    # 不清除合并单元格的非左上角
                    try:
                        cell.value = None
                    except AttributeError:
                        pass

        self.wb.save(self.output_path)
        return self.output_path
