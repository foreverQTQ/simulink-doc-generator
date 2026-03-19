"""
Word 文档生成器 (优化版)
将 Simulink 模型分析结果生成美观的 Word 文档
"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml
from typing import Dict, List, Optional
from datetime import datetime
import os
import re

from .parser import SimulinkModel, SubSystem, Block
from .analyzer import ModelAnalyzer


class DocGenerator:
    """Word 文档生成器（优化版）"""
    
    def __init__(self, model: SimulinkModel, analyzer: ModelAnalyzer):
        self.model = model
        self.analyzer = analyzer
        self.doc = Document()
        
        # 设置文档默认样式
        self._setup_styles()
        
        # 设置页面边距
        sections = self.doc.sections
        for section in sections:
            section.top_margin = Cm(2.5)
            section.bottom_margin = Cm(2.5)
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)
    
    def _setup_styles(self):
        """设置文档样式"""
        # 设置正文样式
        style = self.doc.styles['Normal']
        style.font.name = 'Microsoft YaHei'
        style.font.size = Pt(10.5)
        style._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        
        # 设置标题样式
        for i in range(1, 4):
            style_name = f'Heading {i}'
            if style_name in self.doc.styles:
                style = self.doc.styles[style_name]
                style.font.name = 'Microsoft YaHei'
                style.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)  # 蓝色
                style._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
                
                # 设置不同级别的字号
                if i == 1:
                    style.font.size = Pt(16)
                elif i == 2:
                    style.font.size = Pt(14)
                else:
                    style.font.size = Pt(12)
    
    def _set_table_style(self, table):
        """设置表格样式"""
        # 设置表格边框
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
        
        # 设置表格宽度自适应
        table.autofit = True
        
        # 设置表头样式
        if len(table.rows) > 0:
            for cell in table.rows[0].cells:
                # 设置背景色为浅蓝色
                shading = OxmlElement('w:shd')
                shading.set(qn('w:fill'), 'D9E2F3')
                cell._tc.get_or_add_tcPr().append(shading)
                
                # 设置字体加粗
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.bold = True
    
    def _add_table_row(self, table, cells: List[str], header: bool = False):
        """添加表格行"""
        row = table.add_row()
        for i, text in enumerate(cells):
            row.cells[i].text = text
        return row
    
    def generate(self, output_path: str, options: Dict = None):
        """生成 Word 文档"""
        options = options or {}
        
        # 1. 封面
        self._add_cover()
        
        # 2. 目录占位
        self._add_toc_placeholder()
        
        # 3. 模型概述
        self._add_overview()
        
        # 4. 系统架构
        self._add_architecture()
        
        # 5. 子系统详细说明
        self._add_subsystem_details(options)
        
        # 6. 附录
        if options.get('include_data_dict', True):
            self._add_appendix()
        
        # 保存文档
        self.doc.save(output_path)
    
    def _add_cover(self):
        """添加封面"""
        # 添加空行
        for _ in range(3):
            self.doc.add_paragraph()
        
        # 主标题
        title = self.doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run(self.model.name)
        run.font.size = Pt(26)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
        
        # 副标题
        subtitle = self.doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = subtitle.add_run("模型设计文档")
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(0x5B, 0x5B, 0x5B)
        
        # 添加空行
        for _ in range(4):
            self.doc.add_paragraph()
        
        # 信息表格
        info_table = self.doc.add_table(rows=6, cols=2)
        info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        info_data = [
            ('模型名称', self.model.name),
            ('版本', self.model.version or 'N/A'),
            ('作者', self.model.author or 'N/A'),
            ('创建时间', self._format_date(self.model.created)),
            ('最后修改', self._format_date(self.model.modified)),
            ('文档生成时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        ]
        
        for i, (label, value) in enumerate(info_data):
            info_table.rows[i].cells[0].text = label
            info_table.rows[i].cells[1].text = value
            # 设置第一列加粗
            for paragraph in info_table.rows[i].cells[0].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
        
        # 分页
        self.doc.add_page_break()
    
    def _format_date(self, date_str: str) -> str:
        """格式化日期"""
        if not date_str:
            return 'N/A'
        try:
            # ISO 格式转换
            if 'T' in date_str:
                dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                return dt.strftime('%Y-%m-%d %H:%M')
            return date_str
        except:
            return date_str
    
    def _add_toc_placeholder(self):
        """添加目录占位"""
        self.doc.add_heading('目录', level=1)
        
        note = self.doc.add_paragraph()
        note.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = note.add_run('（在 Word 中右键点击此处，选择"更新域"生成目录）')
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        
        self.doc.add_page_break()
    
    def _add_overview(self):
        """添加模型概述"""
        self.doc.add_heading('1. 模型概述', level=1)
        
        overview = self.analyzer.get_model_overview()
        
        # 1.1 基本信息
        self.doc.add_heading('1.1 基本信息', level=2)
        
        info_table = self.doc.add_table(rows=1, cols=4)
        self._set_table_style(info_table)
        
        # 表头
        headers = ['属性', '值', '属性', '值']
        for i, h in enumerate(headers):
            info_table.rows[0].cells[i].text = h
        
        # 数据（两列布局）
        data = [
            ('模型名称', overview['name'], '模块总数', str(overview['total_blocks'])),
            ('版本', overview['version'] or 'N/A', '子系统数量', str(overview['total_subsystems'])),
            ('作者', overview['author'] or 'N/A', '最大层级深度', str(overview['max_depth'])),
            ('创建时间', self._format_date(self.model.created), '修改时间', self._format_date(self.model.modified)),
        ]
        
        for row_data in data:
            row = info_table.add_row()
            for i, val in enumerate(row_data):
                row.cells[i].text = val
        
        # 1.2 模块类型统计
        self.doc.add_heading('1.2 模块类型统计', level=2)
        
        if overview['block_type_stats']:
            type_table = self.doc.add_table(rows=1, cols=3)
            self._set_table_style(type_table)
            
            # 表头
            headers = ['模块类型', '数量', '示例模块']
            for i, h in enumerate(headers):
                type_table.rows[0].cells[i].text = h
            
            # 数据（只显示前20种类型）
            for stat in overview['block_type_stats'][:20]:
                row = type_table.add_row()
                row.cells[0].text = stat.block_type
                row.cells[1].text = str(stat.count)
                row.cells[2].text = ', '.join(stat.examples[:3])
        
        self.doc.add_page_break()
    
    def _add_architecture(self):
        """添加系统架构"""
        self.doc.add_heading('2. 系统架构', level=1)
        
        # 2.1 层级结构图
        self.doc.add_heading('2.1 层级结构', level=2)
        
        # 用表格展示层级结构
        subsystems = self.analyzer.get_all_subsystems()
        
        hierarchy_table = self.doc.add_table(rows=1, cols=4)
        self._set_table_style(hierarchy_table)
        
        # 表头
        headers = ['层级', '子系统名称', '模块数', '端口(I/O)']
        for i, h in enumerate(headers):
            hierarchy_table.rows[0].cells[i].text = h
        
        # 数据
        for level, sub in subsystems[:50]:  # 只显示前50个
            row = hierarchy_table.add_row()
            row.cells[0].text = str(level)
            row.cells[1].text = '  ' * level + sub.name
            row.cells[2].text = str(len(sub.blocks))
            row.cells[3].text = f'{len(sub.inports)}/{len(sub.outports)}'
        
        if len(subsystems) > 50:
            row = hierarchy_table.add_row()
            row.cells[0].text = '...'
            row.cells[1].text = f'(还有 {len(subsystems) - 50} 个子系统)'
            row.cells[2].text = ''
            row.cells[3].text = ''
        
        # 2.2 子系统概览表
        self.doc.add_heading('2.2 子系统概览', level=2)
        
        summary_table = self.doc.add_table(rows=1, cols=5)
        self._set_table_style(summary_table)
        
        headers = ['子系统名称', '层级', '模块数', '输入端口', '输出端口']
        for i, h in enumerate(headers):
            summary_table.rows[0].cells[i].text = h
        
        for level, sub in subsystems[:30]:
            row = summary_table.add_row()
            row.cells[0].text = sub.name
            row.cells[1].text = str(level)
            row.cells[2].text = str(len(sub.blocks))
            row.cells[3].text = str(len(sub.inports))
            row.cells[4].text = str(len(sub.outports))
        
        self.doc.add_page_break()
    
    def _add_subsystem_details(self, options: Dict):
        """添加子系统详细说明"""
        self.doc.add_heading('3. 子系统详细说明', level=1)
        
        subsystems = self.analyzer.get_all_subsystems()
        
        for i, (level, subsystem) in enumerate(subsystems):
            # 检查是否需要新建页
            if i > 0 and i % 10 == 0:
                self.doc.add_page_break()
            
            # 子系统标题
            title_prefix = '  ' * min(level, 2)  # 最多缩进2级
            self.doc.add_heading(f'{title_prefix}3.{i+1} {subsystem.name}', level=2)
            
            # 获取摘要
            summary = self.analyzer.get_subsystem_summary(subsystem)
            
            # 1. 功能描述（智能生成）
            self.doc.add_heading('功能描述', level=3)
            func_desc = self._generate_function_description(subsystem)
            self.doc.add_paragraph(func_desc)
            
            # 2. 基本信息（表格）
            self.doc.add_heading('基本信息', level=3)
            
            info_table = self.doc.add_table(rows=2, cols=4)
            self._set_table_style(info_table)
            
            info_table.rows[0].cells[0].text = '层级'
            info_table.rows[0].cells[1].text = str(level)
            info_table.rows[0].cells[2].text = '模块数'
            info_table.rows[0].cells[3].text = str(len(subsystem.blocks))
            
            info_table.rows[1].cells[0].text = '输入端口'
            info_table.rows[1].cells[1].text = str(len(subsystem.inports))
            info_table.rows[1].cells[2].text = '输出端口'
            info_table.rows[1].cells[3].text = str(len(subsystem.outports))
            
            # 3. 输入端口表
            if summary['inports']:
                self.doc.add_heading('输入端口', level=3)
                self._add_port_table(summary['inports'])
            
            # 4. 输出端口表
            if summary['outports']:
                self.doc.add_heading('输出端口', level=3)
                self._add_port_table(summary['outports'])
            
            # 5. 内部模块清单（表格）
            if options.get('include_block_list', True) and subsystem.blocks:
                self.doc.add_heading('内部模块', level=3)
                self._add_block_table(subsystem)
            
            # 添加分隔
            self.doc.add_paragraph()
    
    def _generate_function_description(self, subsystem: SubSystem) -> str:
        """
        根据子系统信息智能生成功能描述
        
        从以下方面推断：
        1. 子系统名称 - 通常包含功能关键词
        2. 内部模块类型 - 判断是数据处理、信号转换还是控制逻辑
        3. 端口信息 - 输入输出的数量和名称
        4. 子系统层级 - 父子系统的上下文
        """
        descriptions = []
        
        # 1. 从名称推断
        name_hints = self._analyze_name(subsystem.name)
        if name_hints:
            descriptions.append(name_hints)
        
        # 2. 从模块组成推断
        block_hints = self._analyze_blocks(subsystem)
        if block_hints:
            descriptions.append(block_hints)
        
        # 3. 从端口信息推断
        port_hints = self._analyze_ports(subsystem)
        if port_hints:
            descriptions.append(port_hints)
        
        if descriptions:
            return '，'.join(descriptions) + '。'
        
        return f'该子系统名为"{subsystem.name}"，包含 {len(subsystem.blocks)} 个内部模块。'
    
    def _analyze_name(self, name: str) -> str:
        """从名称推断功能"""
        name_lower = name.lower()
        
        # 常见关键词映射
        keywords = {
            'input': '负责信号输入处理',
            'output': '负责信号输出处理',
            'can': 'CAN总线通信相关功能',
            'adc': 'ADC模数转换信号处理',
            'calculation': '执行计算功能',
            'control': '控制逻辑功能',
            'function': '功能处理模块',
            'signal': '信号处理',
            'data': '数据处理',
            'communication': '通信功能',
            'manager': '管理功能',
            'handler': '处理功能',
            'process': '处理流程',
            'monitor': '监控功能',
            'detect': '检测功能',
            'protect': '保护功能',
            'diagnostic': '诊断功能',
            'bms': '电池管理系统相关',
            'obc': '车载充电机相关',
            'dc_dc': 'DC-DC转换器相关',
            'mcu': '电机控制器相关',
            'gsm': 'GSM通信相关',
            'startup': '启动流程',
            'shutdown': '关闭流程',
            'init': '初始化功能',
            'main': '主控模块',
            'subsystem': '子系统',
        }
        
        for key, desc in keywords.items():
            if key in name_lower:
                return desc
        
        # 如果没有匹配，根据驼峰命名拆分
        words = re.findall('[A-Z][^A-Z]*', name)
        if words:
            return f"涉及{''.join(words)}相关功能"
        
        return ""
    
    def _analyze_blocks(self, subsystem: SubSystem) -> str:
        """从模块组成推断功能"""
        block_types = {}
        for block in subsystem.blocks:
            block_types[block.block_type] = block_types.get(block.block_type, 0) + 1
        
        hints = []
        
        # 根据模块类型推断
        if 'Inport' in block_types:
            hints.append(f"有{block_types['Inport']}个输入端口")
        if 'Outport' in block_types:
            hints.append(f"有{block_types['Outport']}个输出端口")
        if 'SubSystem' in block_types:
            hints.append(f"包含{block_types['SubSystem']}个子系统")
        if 'Gain' in block_types:
            hints.append("包含增益运算")
        if 'Sum' in block_types:
            hints.append("包含求和运算")
        if 'Product' in block_types:
            hints.append("包含乘法运算")
        if 'Logic' in block_types:
            hints.append("包含逻辑运算")
        if 'Switch' in block_types:
            hints.append("包含条件切换")
        if 'Delay' in block_types or 'UnitDelay' in block_types:
            hints.append("包含延时处理")
        if 'BusSelector' in block_types:
            hints.append("包含总线信号选择")
        if 'BusCreator' in block_types:
            hints.append("包含总线信号创建")
        if 'Constant' in block_types:
            hints.append("包含常量定义")
        if 'Reference' in block_types:
            hints.append("引用外部模块")
        
        if hints:
            return '，'.join(hints[:3])  # 最多3个提示
        return ""
    
    def _analyze_ports(self, subsystem: SubSystem) -> str:
        """从端口信息推断功能"""
        if not subsystem.inports and not subsystem.outports:
            return ""
        
        in_count = len(subsystem.inports)
        out_count = len(subsystem.outports)
        
        if in_count > 0 and out_count > 0:
            return f"具有{in_count}路输入和{out_count}路输出"
        elif in_count > 0:
            return f"具有{in_count}路输入"
        elif out_count > 0:
            return f"具有{out_count}路输出"
        
        return ""
    
    def _add_port_table(self, ports: List[Dict]):
        """添加端口表格"""
        if not ports:
            return
        
        table = self.doc.add_table(rows=1, cols=4)
        self._set_table_style(table)
        
        # 表头
        headers = ['端口名称', '端口号', '数据类型', '描述']
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
        
        # 数据
        for port in ports:
            row = table.add_row()
            row.cells[0].text = port['name']
            row.cells[1].text = str(port['number'])
            row.cells[2].text = port['data_type'] or 'N/A'
            row.cells[3].text = port['description'] or '-'
    
    def _add_block_table(self, subsystem: SubSystem):
        """添加模块清单表格"""
        if not subsystem.blocks:
            return
        
        # 按类型分组统计
        blocks_by_type = {}
        for block in subsystem.blocks:
            if block.block_type not in blocks_by_type:
                blocks_by_type[block.block_type] = []
            blocks_by_type[block.block_type].append(block)
        
        table = self.doc.add_table(rows=1, cols=4)
        self._set_table_style(table)
        
        # 表头
        headers = ['模块类型', '数量', '示例', '说明']
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
        
        # 模块类型说明
        type_descriptions = {
            'Inport': '输入端口',
            'Outport': '输出端口',
            'SubSystem': '子系统',
            'Gain': '增益模块',
            'Sum': '求和模块',
            'Product': '乘法模块',
            'Logic': '逻辑运算',
            'Switch': '切换开关',
            'Delay': '延时模块',
            'UnitDelay': '单位延时',
            'BusSelector': '总线选择器',
            'BusCreator': '总线创建器',
            'Constant': '常量',
            'Reference': '引用模块',
            'DataTypeConversion': '数据类型转换',
            'Demux': '解复用器',
            'Mux': '复用器',
            'Memory': '存储器',
            'TriggerPort': '触发端口',
            'Abs': '绝对值',
            'MinMax': '最大最小值',
            'RelationalOperator': '关系运算',
            'S-Function': 'S函数',
            'Terminator': '终端',
            'From': '信号来源',
            'Goto': '信号去向',
        }
        
        for block_type, blocks in blocks_by_type.items():
            row = table.add_row()
            row.cells[0].text = block_type
            row.cells[1].text = str(len(blocks))
            row.cells[2].text = ', '.join([b.name for b in blocks[:3]])
            if len(blocks) > 3:
                row.cells[2].text += f' ...等{len(blocks)}个'
            row.cells[3].text = type_descriptions.get(block_type, '-')
    
    def _add_appendix(self):
        """添加附录"""
        self.doc.add_page_break()
        self.doc.add_heading('附录', level=1)
        
        # A. 数据字典
        self.doc.add_heading('A. 数据字典', level=2)
        
        # 收集所有端口
        all_inports = []
        all_outports = []
        
        for level, subsystem in self.analyzer.get_all_subsystems():
            for port in subsystem.inports:
                all_inports.append({
                    'subsystem': subsystem.name,
                    'level': level,
                    'port': port
                })
            for port in subsystem.outports:
                all_outports.append({
                    'subsystem': subsystem.name,
                    'level': level,
                    'port': port
                })
        
        # A.1 输入端口字典
        if all_inports:
            self.doc.add_heading('A.1 输入端口汇总', level=3)
            
            inport_table = self.doc.add_table(rows=1, cols=5)
            self._set_table_style(inport_table)
            
            headers = ['所属子系统', '端口名称', '端口号', '数据类型', '描述']
            for i, h in enumerate(headers):
                inport_table.rows[0].cells[i].text = h
            
            for item in all_inports[:100]:  # 只显示前100个
                row = inport_table.add_row()
                row.cells[0].text = item['subsystem']
                row.cells[1].text = item['port'].name
                row.cells[2].text = str(item['port'].port_number)
                row.cells[3].text = item['port'].data_type or 'N/A'
                row.cells[4].text = item['port'].description or '-'
            
            if len(all_inports) > 100:
                row = inport_table.add_row()
                row.cells[0].text = '...'
                row.cells[1].text = f'(还有 {len(all_inports) - 100} 个端口)'
        
        # A.2 输出端口字典
        if all_outports:
            self.doc.add_heading('A.2 输出端口汇总', level=3)
            
            outport_table = self.doc.add_table(rows=1, cols=5)
            self._set_table_style(outport_table)
            
            headers = ['所属子系统', '端口名称', '端口号', '数据类型', '描述']
            for i, h in enumerate(headers):
                outport_table.rows[0].cells[i].text = h
            
            for item in all_outports[:100]:
                row = outport_table.add_row()
                row.cells[0].text = item['subsystem']
                row.cells[1].text = item['port'].name
                row.cells[2].text = str(item['port'].port_number)
                row.cells[3].text = item['port'].data_type or 'N/A'
                row.cells[4].text = item['port'].description or '-'
            
            if len(all_outports) > 100:
                row = outport_table.add_row()
                row.cells[0].text = '...'
                row.cells[1].text = f'(还有 {len(all_outports) - 100} 个端口)'


def generate_document(model: SimulinkModel, output_path: str, options: Dict = None):
    """便捷函数：生成文档"""
    analyzer = ModelAnalyzer(model)
    generator = DocGenerator(model, analyzer)
    generator.generate(output_path, options)


if __name__ == "__main__":
    print("文档生成器模块（优化版）")