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
import os
import io
import tempfile
import base64
from datetime import datetime

app = Flask(__name__)
CORS(app)

# 数据映射和提取函数
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
            row_values = [str(sheet.cell_value(row_idx, col_idx)) for col_idx in range(sheet.ncols)]

            # 提取基本信息
            if '船名' in row_values[0]:
                data['vessel_name'] = row_values[1] if len(row_values) > 1 else ''
                data['voyage_no'] = row_values[4] if len(row_values) > 4 else ''
            elif '航次' in row_values[0] and 'voyage_no' not in data:
                data['voyage_no'] = row_values[1] if len(row_values) > 1 else ''
            elif '目的港' in row_values[0]:
                data['port_of_discharge'] = row_values[1] if len(row_values) > 1 else ''
            elif '总提单号' in row_values[0]:
                data['master_bl_no'] = row_values[1] if len(row_values) > 1 else ''

            # 提取分票统计数据
            elif '提单号' in row_values and '英文品名' in row_values:
                # 下一行是数据
                if row_idx + 1 < sheet.nrows:
                    data_row = [str(sheet.cell_value(row_idx + 1, col_idx)) for col_idx in range(sheet.ncols)]
                    data['bl_no'] = data_row[0]
                    data['cargo_name'] = data_row[2]
                    data['marks'] = data_row[5] if len(data_row) > 5 else 'N/M'
                    data['packages'] = data_row[9] if len(data_row) > 9 else ''
                    data['package_unit'] = data_row[10] if len(data_row) > 10 else 'CARTONS'
                    data['gross_weight'] = data_row[11] if len(data_row) > 11 else ''
                    data['volume'] = data_row[12] if len(data_row) > 12 else ''

            # 提取按箱统计数据
            elif '箱号' in row_values and '封号' in row_values:
                # 下一行是数据
                if row_idx + 1 < sheet.nrows:
                    data_row = [str(sheet.cell_value(row_idx + 1, col_idx)) for col_idx in range(sheet.ncols)]
                    data['container_no'] = data_row[0]
                    data['seal_no'] = data_row[1] if len(data_row) > 1 else ''
                    data['container_type'] = data_row[2] if len(data_row) > 2 else ''

        return data

    def generate_bl_confirmation(self, output_path, consignor_info=None, consignee_info=None, notify_party_info=None):
        """生成提单确认件"""
        doc = Document()

        # 设置页面边距
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)

        # 设置中文字体
        def set_chinese_font(run, font_name='宋体', font_size=11):
            run.font.name = font_name
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            run.font.size = Pt(font_size)

        # 船名航次和目的港
        p = doc.add_paragraph()
        run = p.add_run(f"船名航次：{self.manifest_data.get('vessel_name', '')} {self.manifest_data.get('voyage_no', '')}\t")
        set_chinese_font(run)
        run = p.add_run(f"目的港：{self.manifest_data.get('port_of_discharge', '')}\n")
        set_chinese_font(run)

        # 提单号
        p = doc.add_paragraph()
        run = p.add_run(f"提单号：{self.manifest_data.get('bl_no', self.manifest_data.get('master_bl_no', ''))}\n")
        set_chinese_font(run)

        # 箱号和封号
        p = doc.add_paragraph()
        run = p.add_run(f"相应箱号 封号\n")
        set_chinese_font(run)
        run = p.add_run(f"箱号：{self.manifest_data.get('container_no', '')}\n")
        set_chinese_font(run)
        run = p.add_run(f"封号：{self.manifest_data.get('seal_no', '')}\n")
        set_chinese_font(run)
        run = p.add_run(f"箱型：{self.manifest_data.get('container_type', '')}\n")
        set_chinese_font(run)

        # 发货人
        p = doc.add_paragraph()
        run = p.add_run("发货人：\n")
        set_chinese_font(run)
        if consignor_info:
            for line in consignor_info.split('\n'):
                run = p.add_run(f"{line}\n")
                set_chinese_font(run)

        # 收货人
        p = doc.add_paragraph()
        run = p.add_run("收货人：\n")
        set_chinese_font(run)
        if consignee_info:
            for line in consignee_info.split('\n'):
                run = p.add_run(f"{line}\n")
                set_chinese_font(run)

        # 通知人
        p = doc.add_paragraph()
        run = p.add_run("通知人：\n")
        set_chinese_font(run)
        if notify_party_info:
            for line in notify_party_info.split('\n'):
                run = p.add_run(f"{line}\n")
                set_chinese_font(run)

        # 品名
        p = doc.add_paragraph()
        run = p.add_run("品名：\n")
        set_chinese_font(run)
        cargo_names = self.manifest_data.get('cargo_name', '').split(',')
        for name in cargo_names:
            run = p.add_run(f"{name.strip()}\n")
            set_chinese_font(run)

        # 件数/重量/体积
        p = doc.add_paragraph()
        run = p.add_run("件数／重量／体积\n")
        set_chinese_font(run)
        run = p.add_run(f"件数：{self.manifest_data.get('packages', '')} {self.manifest_data.get('package_unit', '')}\n")
        set_chinese_font(run)
        run = p.add_run(f"毛重：{self.manifest_data.get('gross_weight', '')} KGS\n")
        set_chinese_font(run)
        run = p.add_run(f"体积：{self.manifest_data.get('volume', '')} CBM\n")
        set_chinese_font(run)

        # 免用箱申请
        p = doc.add_paragraph()
        run = p.add_run("申请14天免用箱显示在提单上")
        set_chinese_font(run)

        doc.save(output_path)
        return output_path

    def generate_packing_list_invoice(self, output_path, invoice_no=None, invoice_date=None, consignee_name=None, items=None):
        """生成装箱单发票"""
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet = workbook.add_sheet('装箱单发票')

        # 设置列宽
        sheet.col(0).width = 256 * 15
        sheet.col(1).width = 256 * 8
        sheet.col(2).width = 256 * 10
        sheet.col(3).width = 256 * 35
        sheet.col(4).width = 256 * 10
        sheet.col(5).width = 256 * 8
        sheet.col(6).width = 256 * 10
        sheet.col(7).width = 256 * 10

        # 定义样式
        title_style = xlwt.XFStyle()
        title_font = xlwt.Font()
        title_font.height = 280  # 14pt
        title_style.font = title_font

        normal_style = xlwt.XFStyle()
        normal_font = xlwt.Font()
        normal_font.height = 220  # 11pt
        normal_style.font = normal_font

        bold_style = xlwt.XFStyle()
        bold_font = xlwt.Font()
        bold_font.height = 220
        bold_font.bold = True
        bold_style.font = bold_font

        # 第1行：空
        pass

        # 第2行：公司名称
        sheet.write(1, 0, '浙江长江国际有限公司', title_style)
        sheet.write_merge(1, 1, 1, 14, '', normal_style)

        # 第3行：英文名称
        sheet.write(2, 0, 'ZHEJIANG CHEUNG KONG INTERNATIONAL LIMITED', normal_style)
        sheet.write_merge(2, 2, 1, 14, '', normal_style)

        # 第4行：发票标题
        sheet.write(3, 4, '发票', bold_style)
        sheet.write(3, 6, '第', normal_style)
        sheet.write(3, 7, invoice_no or 'YWSJ2602044', normal_style)
        sheet.write(3, 8, '号', normal_style)

        # 第5行：No.
        sheet.write(4, 5, 'No.', normal_style)
        sheet.write_merge(4, 4, 6, 8, '………………………………', normal_style)

        # 第6行：INVOICE和日期
        sheet.write(5, 4, 'INVOICE', bold_style)
        sheet.write(5, 6, '日期', normal_style)
        sheet.write(5, 7, invoice_date or datetime.now().strftime('%b.%d.%Y').upper(), normal_style)

        # 第7行：Date占位
        sheet.write(6, 6, 'Date……………………………………', normal_style)

        # 第8行：收货人
        if consignee_name:
            sheet.write(7, 0, consignee_name, normal_style)
        else:
            sheet.write(7, 0, 'TO: SIJI SHIPPING L.L.C', normal_style)
        sheet.write_merge(7, 7, 1, 4, '', normal_style)
        sheet.write(7, 5, '信用证第', normal_style)
        sheet.write_merge(7, 7, 6, 8, '', normal_style)
        sheet.write(7, 9, '号', normal_style)

        # 第9行：To占位
        sheet.write(8, 0, 'To:…………………………………………………………', normal_style)
        sheet.write(8, 5, 'L/C NO:………………………………', normal_style)

        # 第10行：表头
        sheet.write(9, 0, '唛头号码 Marks & Numbers', bold_style)
        sheet.write_merge(9, 9, 1, 2, '', normal_style)
        sheet.write(9, 3, '数量与品名 Quantities and Descriptions', bold_style)
        sheet.write_merge(9, 9, 4, 5, '', normal_style)
        sheet.write(9, 6, '单价 Unit price', bold_style)
        sheet.write_merge(9, 9, 7, 8, '金额 Amount', bold_style)

        # 第11行：N/M
        marks = self.manifest_data.get('marks', 'N/M')
        sheet.write(10, 0, marks, normal_style)
        sheet.write_merge(10, 10, 1, 4, 'CIF DUBAI', normal_style)

        # 商品明细行
        if items:
            row_idx = 11
            for item in items:
                sheet.write(row_idx, 1, str(item.get('qty', '')), normal_style)
                sheet.write(row_idx, 2, item.get('unit', 'CTNS'), normal_style)
                sheet.write(row_idx, 3, item.get('name', ''), normal_style)
                sheet.write(row_idx, 5, str(item.get('unit_price', '')), normal_style)
                sheet.write(row_idx, 6, '/CTNS', normal_style)
                sheet.write(row_idx, 7, str(item.get('amount', '')), normal_style)
                row_idx += 1
        else:
            # 默认示例数据
            sheet.write(11, 1, '14.0')
            sheet.write(11, 2, 'CTNS')
            sheet.write(11, 3, 'PLASTIC CAP')
            sheet.write(11, 5, '11.0')
            sheet.write(11, 6, '/CTNS')
            sheet.write(11, 7, '154.0')

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


@app.route('/api/process', methods=['POST'])
def process_manifest():
    """处理舱单并生成文档"""
    try:
        # 获取上传的文件
        if 'manifest_file' not in request.files:
            return jsonify({'error': '请上传舱单文件'}), 400

        file = request.files['manifest_file']
        if file.filename == '':
            return jsonify({'error': '未选择文件'}), 400

        # 保存临时文件
        temp_manifest = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
        file.save(temp_manifest.name)
        temp_manifest.close()

        # 获取表单数据
        invoice_no = request.form.get('invoice_no', '')
        invoice_date = request.form.get('invoice_date', '')
        consignor = request.form.get('consignor', '')
        consignee = request.form.get('consignee', '')
        notify_party = request.form.get('notify_party', '')

        # 解析商品明细
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

        # 处理舱单
        processor = ManifestProcessor(temp_manifest.name)

        # 生成提单确认件
        bl_output = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        processor.generate_bl_confirmation(
            bl_output.name,
            consignor_info=consignor,
            consignee_info=consignee,
            notify_party_info=notify_party
        )

        # 生成装箱单发票
        pl_output = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
        processor.generate_packing_list_invoice(
            pl_output.name,
            invoice_no=invoice_no,
            invoice_date=invoice_date,
            consignee_name=consignee.split('\n')[0] if consignee else None,
            items=items if items else None
        )

        # 将文件转换为 base64
        bl_data = file_to_base64(bl_output.name)
        pl_data = file_to_base64(pl_output.name)

        # 清理临时文件
        try:
            os.unlink(temp_manifest.name)
            os.unlink(bl_output.name)
            os.unlink(pl_output.name)
        except:
            pass

        # 返回 base64 编码的文件数据
        return jsonify({
            'success': True,
            'message': '文档生成成功！',
            'bl_document': bl_data,
            'pl_document': pl_data,
            'extracted_data': processor.manifest_data
        })

    except Exception as e:
        return jsonify({'error': f'处理失败: {str(e)}'}), 500
    finally:
        # 清理临时文件
        try:
            if 'temp_manifest' in locals():
                os.unlink(temp_manifest.name)
        except:
            pass


@app.route('/api/preview', methods=['POST'])
def preview_data():
    """预览提取的数据"""
    try:
        if 'manifest_file' not in request.files:
            return jsonify({'error': '请上传舱单文件'}), 400

        file = request.files['manifest_file']
        if file.filename == '':
            return jsonify({'error': '未选择文件'}), 400

        # 保存临时文件
        temp_manifest = tempfile.NamedTemporaryFile(delete=False, suffix='.xls')
        file.save(temp_manifest.name)
        temp_manifest.close()

        # 处理舱单
        processor = ManifestProcessor(temp_manifest.name)

        return jsonify({
            'success': True,
            'data': processor.manifest_data
        })

    except Exception as e:
        return jsonify({'error': f'预览失败: {str(e)}'}), 500
    finally:
        try:
            if 'temp_manifest' in locals():
                os.unlink(temp_manifest.name)
        except:
            pass


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
