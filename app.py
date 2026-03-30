# -*- coding: utf-8 -*-
"""
舱单数据处理应用
提取舱单数据并生成提单确认件和装箱单发票
"""

from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
import xlrd
import xlwt
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import io
import tempfile
import base64
from datetime import datetime

app = Flask(__name__)
CORS(app)


def set_cell_border(cell, **kwargs):
    """设置单元格边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        if edge in kwargs:
            edge_element = OxmlElement(f'w:{edge}')
            edge_element.set(qn('w:val'), kwargs[edge])
            tcBorders.append(edge_element)

    tcPr.append(tcBorders)
    return cell


class ManifestProcessor:
    """舱单数据处理器"""

    def __init__(self, manifest_file_path):
        """初始化处理器"""
        self.manifest_data = self._read_manifest(manifest_file_path)

    def _read_manifest(self, file_path):
        """读取舱单Excel文件"""
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)

        data = {}
        for row_idx in range(sheet.nrows):
            row_values = [str(sheet.cell_value(row_idx, col_idx)) if sheet.cell_value(row_idx, col_idx) else '' for col_idx in range(sheet.ncols)]

            # 1. 提取基本信息 (Row 3): 船名、航次、目的港
            if '船名' in row_values[0] and 'vessel_name' not in data:
                data['vessel_name'] = row_values[1] if len(row_values) > 1 else ''
                data['voyage_no'] = row_values[4] if len(row_values) > 4 else ''
                # 同时提取目的港（在同一行，查找"目的港"列）
                for col_idx, val in enumerate(row_values):
                    if val == '目的港' and col_idx + 1 < len(row_values) and 'port_of_discharge' not in data:
                        data['port_of_discharge'] = row_values[col_idx + 1]

            # 2. 提取总提单号 (Row 4)
            elif '总提单号' in row_values[0] and 'master_bl_no' not in data:
                data['master_bl_no'] = row_values[1] if len(row_values) > 1 else ''

            # 3. 提取分票统计数据 (Row 7-8) - 提单号、英文品名、唛头、件数、包装单位、毛重、体积
            elif ('提单号' in row_values and '英文品名' in row_values) and 'bl_no' not in data:
                if row_idx + 1 < sheet.nrows:
                    data_row = [str(sheet.cell_value(row_idx + 1, col_idx)) for col_idx in range(sheet.ncols)]
                    data['bl_no'] = data_row[0]
                    data['cargo_name'] = data_row[2]
                    data['marks'] = data_row[6] if len(data_row) > 6 else 'N/M'
                    data['packages'] = data_row[9] if len(data_row) > 9 else ''
                    data['package_unit'] = data_row[10] if len(data_row) > 10 else 'CARTONS'
                    data['gross_weight'] = data_row[11] if len(data_row) > 11 else ''
                    data['volume'] = data_row[12] if len(data_row) > 12 else ''

            # 4. 提取按箱统计数据 (Row 11-12) - 箱号、封号、箱型
            elif ('箱号' in row_values and '封号' in row_values and '提单号' in row_values) and 'container_no' not in data:
                if row_idx + 1 < sheet.nrows:
                    data_row = [str(sheet.cell_value(row_idx + 1, col_idx)) for col_idx in range(sheet.ncols)]
                    data['container_no'] = data_row[0]
                    data['seal_no'] = data_row[1] if len(data_row) > 1 else ''
                    data['container_type'] = data_row[2] if len(data_row) > 2 else ''

            # 5. 提取发货人信息 (Row 26-31)
            elif ('发货人' in row_values[0] or ('发货' in row_values[0] and 'Shipper' in row_values[0])) and 'shipper' not in data:
                shipper = []
                for i in range(row_idx + 2, min(row_idx + 6, sheet.nrows)):
                    row_vals = [str(sheet.cell_value(i, col)) for col in range(sheet.ncols)]
                    if '名称' in row_vals[0] and row_vals[1]:
                        shipper.append(row_vals[1])
                    elif '地址' in row_vals[0] and row_vals[1]:
                        shipper.append(row_vals[1])
                    elif '电话' in row_vals[0] and row_vals[1]:
                        shipper.append(f'TEL: {row_vals[1]}')
                data['shipper'] = '\\n'.join(shipper) if shipper else ''

            # 6. 提取收货人信息 (Row 33-40)
            elif ('收货人' in row_values[0] or ('收货' in row_values[0] and 'Consignee' in row_values[0])) and 'consignee' not in data:
                consignee = []
                for i in range(row_idx + 2, min(row_idx + 8, sheet.nrows)):
                    row_vals = [str(sheet.cell_value(i, col)) for col in range(sheet.ncols)]
                    if '名称' in row_vals[0] and row_vals[1]:
                        consignee.append(row_vals[1])
                    elif '地址' in row_vals[0] and row_vals[1]:
                        consignee.append(row_vals[1])
                    elif '电话' in row_vals[0] and row_vals[1]:
                        consignee.append(f'TEL: {row_vals[1]}')
                    elif '具体联系人' in row_vals[0] and row_vals[1]:
                        consignee.append(f'ATTN: {row_vals[1]}')
                    elif '联系人电话' in row_vals[0] and row_vals[1]:
                        consignee.append(f'MOB: {row_vals[1]}')
                data['consignee'] = '\\n'.join(consignee) if consignee else ''

            # 7. 提取通知人信息 (Row 42-47)
            elif ('通知人' in row_values[0] or ('通知' in row_values[0] and 'Notifier' in row_values[0])) and 'notifier' not in data:
                notifier = []
                for i in range(row_idx + 2, min(row_idx + 6, sheet.nrows)):
                    row_vals = [str(sheet.cell_value(i, col)) for col in range(sheet.ncols)]
                    if '名称' in row_vals[0] and row_vals[1]:
                        notifier.append(row_vals[1])
                    elif '地址' in row_vals[0] and row_vals[1]:
                        notifier.append(row_vals[1])
                    elif '电话' in row_vals[0] and row_vals[1]:
                        notifier.append(f'TEL: {row_vals[1]}')
                data['notifier'] = '\\n'.join(notifier) if notifier else ''

        return data

    def generate_bl_confirmation(self, output_path, consignor_info=None, consignee_info=None, notify_party_info=None):
        """生成提单确认件"""
        doc = Document()

        # 设置页面边距
        for section in doc.sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)

        def add_run(paragraph, text, font_size=11, bold=False):
            run = paragraph.add_run(text)
            run.font.name = 'Arial'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.size = Pt(font_size)
            run.font.bold = bold
            return run

        # 船名航次和目的港
        p = doc.add_paragraph()
        add_run(p, f"船名航次：{self.manifest_data.get('vessel_name', '')} {self.manifest_data.get('voyage_no', '')}\t")
        add_run(p, f"目的港：{self.manifest_data.get('port_of_discharge', '')}\n")

        # 提单号
        p = doc.add_paragraph()
        add_run(p, f"提单号：{self.manifest_data.get('bl_no', self.manifest_data.get('master_bl_no', ''))}\n")

        # 箱号和封号
        p = doc.add_paragraph()
        add_run(p, f"相应箱号 封号\n")
        add_run(p, f"箱号：{self.manifest_data.get('container_no', '')}\n")
        add_run(p, f"封号：{self.manifest_data.get('seal_no', '')}\n")
        add_run(p, f"箱型：{self.manifest_data.get('container_type', '')}\n")

        # 发货人 - 优先使用舱单提取的数据
        p = doc.add_paragraph()
        add_run(p, "发货人：\n")
        shipper = self.manifest_data.get('shipper', '')
        if shipper:
            for line in shipper.split('\\n'):
                add_run(p, f"{line}\n")
        elif consignor_info:
            for line in consignor_info.split('\n'):
                add_run(p, f"{line}\n")

        # 收货人 - 优先使用舱单提取的数据
        p = doc.add_paragraph()
        add_run(p, "收货人：\n")
        consignee = self.manifest_data.get('consignee', '')
        if consignee:
            for line in consignee.split('\\n'):
                add_run(p, f"{line}\n")
        elif consignee_info:
            for line in consignee_info.split('\n'):
                add_run(p, f"{line}\n")

        # 通知人 - 优先使用舱单提取的数据
        p = doc.add_paragraph()
        add_run(p, "通知人：\n")
        notifier = self.manifest_data.get('notifier', '')
        if notifier:
            for line in notifier.split('\\n'):
                add_run(p, f"{line}\n")
        elif notify_party_info:
            for line in notify_party_info.split('\n'):
                add_run(p, f"{line}\n")

        # 品名
        p = doc.add_paragraph()
        add_run(p, "品名：\n")
        cargo_names = self.manifest_data.get('cargo_name', '').split(',')
        for name in cargo_names:
            add_run(p, f"{name.strip()}\n")

        # 件数/重量/体积 - 使用舱单提取的数据
        p = doc.add_paragraph()
        add_run(p, "件数／重量／体积\n")
        add_run(p, f"件数：{self.manifest_data.get('packages', '')} {self.manifest_data.get('package_unit', '')}\n")
        add_run(p, f"毛重：{self.manifest_data.get('gross_weight', '')} KGS\n")
        add_run(p, f"体积：{self.manifest_data.get('volume', '')} CBM\n")

        # 免用箱申请
        p = doc.add_paragraph()
        add_run(p, "申请14天免用箱显示在提单上")

        doc.save(output_path)
        return output_path

    def generate_packing_list_invoice(self, output_path, invoice_no=None, invoice_date=None, consignee_name=None, items=None):
        """生成装箱单发票 - 完全按照格式要求"""
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet = workbook.add_sheet('装箱单发票')

        # 设置列宽 (基于 Excel 列宽单位，1单位约等于1/256字符宽)
        sheet.col(0).width = 256 * 12
        sheet.col(1).width = 256 * 2
        sheet.col(2).width = 256 * 6
        sheet.col(3).width = 256 * 8
        sheet.col(4).width = 256 * 30
        sheet.col(5).width = 256 * 10
        sheet.col(6).width = 256 * 5
        sheet.col(7).width = 256 * 10

        # 定义样式
        def get_style(bold=False, horiz_align='LEFT', font_size=11):
            style = xlwt.XFStyle()
            font = xlwt.Font()
            font.height = 220 * font_size // 11  # 转换点数
            font.bold = bold
            font.name = 'Arial'
            style.font = font

            alignment = xlwt.Alignment()
            if horiz_align == 'CENTER':
                alignment.horz = xlwt.Alignment.HORZ_CENTER
            elif horiz_align == 'RIGHT':
                alignment.horz = xlwt.Alignment.HORZ_RIGHT
            else:
                alignment.horz = xlwt.Alignment.HORZ_LEFT
            style.alignment = alignment

            return style

        def get_border_style(thin=True):
            style = xlwt.XFStyle()
            borders = xlwt.Borders()
            if thin:
                borders.left = xlwt.Borders.THIN
                borders.right = xlwt.Borders.THIN
                borders.top = xlwt.Borders.THIN
                borders.bottom = xlwt.Borders.THIN
            style.borders = borders
            return style

        # 定义带边框的样式
        border_style = get_border_style()
        border_style.font = get_style().font
        border_style.alignment = get_style().alignment

        # Row 0: 空行
        pass

        # Row 1: 公司名称 (合并单元格 A1-O1)
        title_style = get_style(bold=True)
        sheet.write_merge(1, 1, 0, 14, '浙江长江国际有限公司', title_style)

        # Row 2: 英文名称 (合并单元格 A2-O2)
        normal_style = get_style()
        sheet.write_merge(2, 2, 0, 14, '            ZHEJIANG CHEUNG KONG INTERNATIONAL LIMITED', normal_style)

        # Row 3: 发票 第 [发票号] 号
        bold_style = get_style(bold=True)
        sheet.write(3, 4, '发票', bold_style)
        sheet.write(3, 5, '第', normal_style)
        sheet.write(3, 6, invoice_no or 'YWSJ2602044', normal_style)
        sheet.write(3, 7, '号', normal_style)

        # Row 4: No. 和占位符
        sheet.write(4, 5, 'No.', normal_style)
        sheet.write_merge(4, 4, 6, 8, '………………………………', normal_style)

        # Row 5: INVOICE 日期 [日期]
        sheet.write(5, 4, 'INVOICE', bold_style)
        sheet.write(5, 5, '日期', normal_style)
        sheet.write(5, 6, invoice_date or datetime.now().strftime('%b.%d.%Y').upper(), normal_style)

        # Row 6: Date 占位符
        sheet.write(6, 5, 'Date……………………………………', normal_style)

        # Row 7: 收货人 | 信用证第 [ ] 号
        consignee = consignee_name or self.manifest_data.get('consignee', 'SIJI SHIPPING L.L.C')
        # 简化收货人名称（只取第一行或名称部分）
        if '\\n' in consignee:
            consignee = consignee.split('\\n')[0]
        sheet.write(7, 0, consignee, normal_style)
        sheet.write(7, 5, '信用证第', normal_style)
        sheet.write(7, 6, '', normal_style)
        sheet.write(7, 7, '号', normal_style)

        # Row 8: To占位 | L/C NO占位
        sheet.write(8, 0, 'To:…………………………………………………………', normal_style)
        sheet.write(8, 5, 'L/C NO:………………………………', normal_style)

        # Row 9: 表头 (带边框)
        header_style = get_style(bold=True)
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        header_style.borders = borders

        sheet.write(9, 0, '唛头号码    Marks & Numbers', header_style)
        sheet.write_merge(9, 9, 1, 2, '', header_style)
        sheet.write(9, 3, '数量与品名                                  Quantities and Descriptions', header_style)
        sheet.write_merge(9, 9, 4, 5, '', header_style)
        sheet.write(9, 6, '单价            Unit price', header_style)
        sheet.write_merge(9, 9, 7, 8, '金额       Amount', header_style)

        # Row 10: N/M | CIF DUBAI
        marks = self.manifest_data.get('marks', 'N/M')
        sheet.write(10, 0, marks, header_style)
        sheet.write_merge(10, 10, 1, 4, 'CIF DUBAI', header_style)

        # 商品明细行
        if items:
            row_idx = 11
            for item in items:
                sheet.write(row_idx, 2, str(item.get('qty', '')), border_style)
                sheet.write(row_idx, 3, item.get('unit', 'CTNS'), border_style)
                sheet.write(row_idx, 4, item.get('name', ''), border_style)
                sheet.write(row_idx, 5, str(item.get('unit_price', '')), border_style)
                sheet.write(row_idx, 6, '/CTNS', border_style)
                sheet.write(row_idx, 7, str(item.get('amount', '')), border_style)
                row_idx += 1

        workbook.save(output_path)
        return output_path


@app.route('/')
def index():
    """主页"""
    return render_template('index.html')


def file_to_base64(file_path):
    """将文件转换为 base64 编码"""
    with open(file_path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')


@app.route('/api/preview', methods=['POST', 'OPTIONS'])
def preview_data():
    """预览提取的数据"""
    if request.method == 'OPTIONS':
        return jsonify({}), 200

    try:
        if 'manifest_file' not in request.files:
            return jsonify({'error': '请上传舱单文件'}), 400

        file = request.files['manifest_file']
        if file.filename == '':
            return jsonify({'error': '未选择文件'}), 400

        temp_manifest = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
        file.save(temp_manifest.name)
        temp_manifest.close()

        processor = ManifestProcessor(temp_manifest.name)

        try:
            os.unlink(temp_manifest.name)
        except:
            pass

        return jsonify({
            'success': True,
            'data': processor.manifest_data
        })

    except Exception as e:
        return jsonify({'error': f'预览失败: {str(e)}'}), 500


@app.route('/api/process', methods=['POST', 'OPTIONS'])
def process_manifest():
    """处理舱单并生成文档"""
    if request.method == 'OPTIONS':
        return jsonify({}), 200

    try:
        if 'manifest_file' not in request.files:
            return jsonify({'error': '请上传舱单文件'}), 400

        file = request.files['manifest_file']
        if file.filename == '':
            return jsonify({'error': '未选择文件'}), 400

        temp_manifest = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
        file.save(temp_manifest.name)
        temp_manifest.close()

        invoice_no = request.form.get('invoice_no', '')
        invoice_date = request.form.get('invoice_date', '')
        consignor = request.form.get('consignor', '')
        consignee = request.form.get('consignee', '')
        notify_party = request.form.get('notify_party', '')

        items = []
        items_data = request.form.get('items', '')
        if items_data:
            for item_line in items_data.split('\n'):
                parts = item_line.split('|')
                if len(parts) >= 4:
                    items.append({
                        'qty': parts[0].strip(),
                        'unit': parts[1].strip(),
                        'name': parts[2].strip(),
                        'unit_price': parts[3].strip(),
                        'amount': parts[4].strip() if len(parts) > 4 else ''
                    })

        processor = ManifestProcessor(temp_manifest.name)

        bl_output = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        processor.generate_bl_confirmation(
            bl_output.name,
            consignor_info=consignor,
            consignee_info=consignee,
            notify_party_info=notify_party
        )

        pl_output = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
        processor.generate_packing_list_invoice(
            pl_output.name,
            invoice_no=invoice_no,
            invoice_date=invoice_date,
            consignee_name=consignee.split('\n')[0] if consignee else None,
            items=items if items else None
        )

        bl_data = file_to_base64(bl_output.name)
        pl_data = file_to_base64(pl_output.name)

        try:
            os.unlink(temp_manifest.name)
            os.unlink(bl_output.name)
            os.unlink(pl_output.name)
        except:
            pass

        return jsonify({
            'success': True,
            'message': '文档生成成功！',
            'bl_document': bl_data,
            'pl_document': pl_data,
            'extracted_data': processor.manifest_data
        })

    except Exception as e:
        import traceback
        return jsonify({'error': f'处理失败: {str(e)}\n{traceback.format_exc()}'}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
