"""
Simulink SLX 文件解析器（增强版）
解析 .slx 文件，提取模型结构信息和信号线连接关系
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field
import re


@dataclass
class Port:
    """端口信息"""
    name: str
    port_type: str
    port_number: int
    data_type: str = ""
    dimensions: str = ""
    description: str = ""
    

@dataclass
class Block:
    """模块信息"""
    sid: str
    name: str
    block_type: str
    parent: str = ""
    position: tuple = (0, 0, 0, 0)
    parameters: Dict = field(default_factory=dict)
    ports: List[Dict] = field(default_factory=list)
    description: str = ""
    

@dataclass
class SignalLine:
    """信号线"""
    src_sid: str
    src_port: int
    dst_sid: str
    dst_port: int
    name: str = ""
    

@dataclass
class SubSystem:
    """子系统信息"""
    sid: str
    name: str
    parent: str = ""
    blocks: List[Block] = field(default_factory=list)
    inports: List[Port] = field(default_factory=list)
    outports: List[Port] = field(default_factory=list)
    signals: List[SignalLine] = field(default_factory=list)
    description: str = ""
    children: List['SubSystem'] = field(default_factory=list)
    
    # 信号流图：sid -> [(dst_sid, src_port, dst_port)]
    signal_graph: Dict[str, List[Tuple[str, int, int]]] = field(default_factory=dict)
    
    # 反向信号图：sid <- [(src_sid, src_port, dst_port)]
    reverse_signal_graph: Dict[str, List[Tuple[str, int, int]]] = field(default_factory=dict)


@dataclass
class SimulinkModel:
    """Simulink 模型"""
    name: str
    version: str = ""
    created: str = ""
    modified: str = ""
    author: str = ""
    description: str = ""
    root_system: Optional[SubSystem] = None
    all_subsystems: Dict[str, SubSystem] = field(default_factory=dict)
    
    # 所有模块的字典，方便查找
    all_blocks: Dict[str, Block] = field(default_factory=dict)


class SLXParser:
    """SLX 文件解析器"""
    
    def __init__(self, slx_path: str):
        self.slx_path = Path(slx_path)
        self.model = None
        
    def parse(self) -> SimulinkModel:
        """解析 SLX 文件"""
        if not self.slx_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.slx_path}")
            
        if not zipfile.is_zipfile(self.slx_path):
            raise ValueError(f"不是有效的 SLX 文件: {self.slx_path}")
            
        with zipfile.ZipFile(self.slx_path, 'r') as zf:
            self.model = self._parse_metadata(zf)
            self._parse_blockdiagram(zf)
            
        return self.model
    
    def _parse_metadata(self, zf: zipfile.ZipFile) -> SimulinkModel:
        """解析元信息"""
        model_name = self.slx_path.stem
        version = ""
        created = ""
        modified = ""
        author = ""
        
        try:
            core_xml = zf.read('metadata/coreProperties.xml')
            root = ET.fromstring(core_xml)
            
            ns = {
                'dc': 'http://purl.org/dc/elements/1.1/',
                'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                'dcterms': 'http://purl.org/dc/terms/'
            }
            
            created_elem = root.find('.//dcterms:created', ns)
            if created_elem is not None:
                created = created_elem.text or ""
                
            modified_elem = root.find('.//dcterms:modified', ns)
            if modified_elem is not None:
                modified = modified_elem.text or ""
                
            creator_elem = root.find('.//dc:creator', ns)
            if creator_elem is not None:
                author = creator_elem.text or ""
                
        except KeyError:
            pass
            
        return SimulinkModel(
            name=model_name,
            version=version,
            created=created,
            modified=modified,
            author=author
        )
    
    def _parse_blockdiagram(self, zf: zipfile.ZipFile):
        """解析 blockdiagram.xml"""
        try:
            xml_content = zf.read('simulink/blockdiagram.xml')
        except KeyError:
            raise ValueError("无法找到 simulink/blockdiagram.xml 文件")
            
        root = ET.fromstring(xml_content)
        
        model_elem = root.find('.//Model')
        if model_elem is None:
            raise ValueError("无法找到 Model 节点")
            
        version_elem = model_elem.find('.//P[@Name="ComputedModelVersion"]')
        if version_elem is not None:
            self.model.version = version_elem.text or ""
        
        system_elem = model_elem.find('.//System')
        if system_elem is None:
            raise ValueError("无法找到根 System 节点")
            
        self.model.root_system = self._parse_system(system_elem, parent_sid="", system_name="Root")
        
        self._build_subsystem_dict(self.model.root_system)
        
    def _parse_system(self, elem: ET.Element, parent_sid: str = "", system_name: str = "Root") -> SubSystem:
        """解析 System 节点"""
        subsystem = SubSystem(
            sid=parent_sid,
            name=system_name,
            parent=parent_sid
        )
        
        # 解析 Block
        for block_elem in elem.findall('Block'):
            block = self._parse_block(block_elem)
            subsystem.blocks.append(block)
            self.model.all_blocks[block.sid] = block
            
            if block.block_type == 'Inport':
                port = self._create_port_from_block(block, 'inport')
                subsystem.inports.append(port)
            elif block.block_type == 'Outport':
                port = self._create_port_from_block(block, 'outport')
                subsystem.outports.append(port)
                
            if block.block_type == 'SubSystem':
                inner_system_elem = block_elem.find('System')
                if inner_system_elem is not None:
                    child_subsystem = self._parse_system(
                        inner_system_elem, 
                        parent_sid=block.sid,
                        system_name=block.name
                    )
                    subsystem.children.append(child_subsystem)
        
        # 解析信号线
        self._parse_signals(elem, subsystem)
        
        return subsystem
    
    def _parse_signals(self, elem: ET.Element, subsystem: SubSystem):
        """解析信号线"""
        # 解析 Line 节点
        for line_elem in elem.findall('Line'):
            self._parse_line_element(line_elem, subsystem)
        
        # 构建信号流图
        self._build_signal_graph(subsystem)
    
    def _parse_line_element(self, line_elem: ET.Element, subsystem: SubSystem):
        """解析单个 Line 节点"""
        src = ""
        dst = ""
        
        # 查找 Src 和 Dst
        for p_elem in line_elem.findall('P'):
            name = p_elem.get('Name', '')
            if name == 'Src':
                src = p_elem.text or ''
            elif name == 'Dst':
                dst = p_elem.text or ''
        
        # 解析格式: SID#out:port 或 SID#in:port
        if src and dst:
            src_sid, src_port = self._parse_signal_endpoint(src, is_output=True)
            dst_sid, dst_port = self._parse_signal_endpoint(dst, is_output=False)
            
            if src_sid and dst_sid:
                signal = SignalLine(
                    src_sid=src_sid,
                    src_port=src_port,
                    dst_sid=dst_sid,
                    dst_port=dst_port
                )
                subsystem.signals.append(signal)
        
        # 递归解析 Branch
        for branch_elem in line_elem.findall('Branch'):
            self._parse_branch_element(branch_elem, subsystem, src)
    
    def _parse_branch_element(self, branch_elem: ET.Element, subsystem: SubSystem, parent_src: str):
        """解析 Branch 节点"""
        dst = ""
        
        for p_elem in branch_elem.findall('P'):
            name = p_elem.get('Name', '')
            if name == 'Dst':
                dst = p_elem.text or ''
        
        if parent_src and dst:
            src_sid, src_port = self._parse_signal_endpoint(parent_src, is_output=True)
            dst_sid, dst_port = self._parse_signal_endpoint(dst, is_output=False)
            
            if src_sid and dst_sid:
                signal = SignalLine(
                    src_sid=src_sid,
                    src_port=src_port,
                    dst_sid=dst_sid,
                    dst_port=dst_port
                )
                subsystem.signals.append(signal)
        
        # 递归解析嵌套的 Branch
        for nested_branch in branch_elem.findall('Branch'):
            self._parse_branch_element(nested_branch, subsystem, parent_src)
    
    def _parse_signal_endpoint(self, endpoint: str, is_output: bool) -> Tuple[str, int]:
        """解析信号端点，返回 (SID, port_number)"""
        # 格式: SID#out:port 或 SID#in:port
        try:
            if '#' in endpoint:
                sid_part, port_part = endpoint.split('#')
                if ':' in port_part:
                    _, port_str = port_part.split(':')
                    port_num = int(port_str)
                    return sid_part, port_num
        except:
            pass
        return "", 1
    
    def _build_signal_graph(self, subsystem: SubSystem):
        """构建信号流图"""
        for signal in subsystem.signals:
            # 正向图：src -> [dst]
            if signal.src_sid not in subsystem.signal_graph:
                subsystem.signal_graph[signal.src_sid] = []
            subsystem.signal_graph[signal.src_sid].append(
                (signal.dst_sid, signal.src_port, signal.dst_port)
            )
            
            # 反向图：dst <- [src]
            if signal.dst_sid not in subsystem.reverse_signal_graph:
                subsystem.reverse_signal_graph[signal.dst_sid] = []
            subsystem.reverse_signal_graph[signal.dst_sid].append(
                (signal.src_sid, signal.src_port, signal.dst_port)
            )
    
    def _parse_block(self, elem: ET.Element) -> Block:
        """解析 Block 节点"""
        sid = elem.get('SID', '')
        name = elem.get('Name', '')
        block_type = elem.get('BlockType', 'Unknown')
        
        parameters = {}
        for p_elem in elem.findall('P'):
            param_name = p_elem.get('Name', '')
            param_value = p_elem.text or ''
            if param_name:
                parameters[param_name] = param_value
        
        position = (0, 0, 0, 0)
        pos_str = parameters.get('Position', '')
        if pos_str:
            try:
                coords = eval(pos_str)
                if isinstance(coords, (list, tuple)) and len(coords) >= 4:
                    position = tuple(coords[:4])
            except:
                pass
        
        ports = []
        for port_elem in elem.findall('Port'):
            port_info = {}
            for p_elem in port_elem.findall('P'):
                port_info[p_elem.get('Name', '')] = p_elem.text or ''
            ports.append(port_info)
        
        description = parameters.get('Description', '')
        
        return Block(
            sid=sid,
            name=name,
            block_type=block_type,
            position=position,
            parameters=parameters,
            ports=ports,
            description=description
        )
    
    def _create_port_from_block(self, block: Block, port_type: str) -> Port:
        """从 Block 创建 Port 对象"""
        port_number = 1
        if 'Port' in block.parameters:
            try:
                port_number = int(block.parameters['Port'])
            except:
                pass
        
        if block.ports:
            port_info = block.ports[0]
            if 'PortNumber' in port_info:
                try:
                    port_number = int(port_info['PortNumber'])
                except:
                    pass
            port_name = port_info.get('Name', block.name)
        else:
            port_name = block.name
        
        data_type = block.parameters.get('OutDataTypeStr', '')
        
        return Port(
            name=port_name,
            port_type=port_type,
            port_number=port_number,
            data_type=data_type,
            description=block.description
        )
    
    def _build_subsystem_dict(self, subsystem: SubSystem):
        """构建子系统字典"""
        if subsystem.sid:
            self.model.all_subsystems[subsystem.sid] = subsystem
            
        for child in subsystem.children:
            self._build_subsystem_dict(child)


def parse_slx(slx_path: str) -> SimulinkModel:
    """便捷函数：解析 SLX 文件"""
    parser = SLXParser(slx_path)
    return parser.parse()


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("用法: python parser.py <slx_file>")
        sys.exit(1)
        
    try:
        model = parse_slx(sys.argv[1])
        print(f"模型名称: {model.name}")
        print(f"版本: {model.version}")
        print(f"创建时间: {model.created}")
        print(f"修改时间: {model.modified}")
        print(f"作者: {model.author}")
        
        if model.root_system:
            print(f"\n根系统模块数: {len(model.root_system.blocks)}")
            print(f"信号线数: {len(model.root_system.signals)}")
            print(f"子系统数: {len(model.root_system.children)}")
            
            # 显示信号流
            if model.root_system.signals:
                print("\n信号连接示例:")
                for signal in model.root_system.signals[:5]:
                    src_block = model.all_blocks.get(signal.src_sid)
                    dst_block = model.all_blocks.get(signal.dst_sid)
                    src_name = src_block.name if src_block else signal.src_sid
                    dst_name = dst_block.name if dst_block else signal.dst_sid
                    print(f"  {src_name} -> {dst_name}")
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()