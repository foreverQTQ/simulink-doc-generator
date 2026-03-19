"""
Simulink SLX 文件解析器
解析 .slx 文件（本质是 ZIP + XML），提取模型结构信息

适配实际 XML 格式：
- 文件路径: simulink/blockdiagram.xml
- 根节点: <ModelInformation><Model>
- 系统: <System>
- 模块: <Block BlockType="..." Name="..." SID="...">
- 参数: <P Name="...">value</P>
- 端口: <Port><P Name="PortNumber">...</P></Port>
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field
import re


@dataclass
class Port:
    """端口信息"""
    name: str
    port_type: str  # 'inport' 或 'outport'
    port_number: int
    data_type: str = ""
    dimensions: str = ""
    description: str = ""
    

@dataclass
class Block:
    """模块信息"""
    sid: str                    # Simulink ID
    name: str                   # 模块名称
    block_type: str             # 模块类型 (Inport, Outport, SubSystem, Gain, etc.)
    parent: str = ""            # 父系统 SID
    position: tuple = (0, 0, 0, 0)  # 位置坐标
    parameters: Dict = field(default_factory=dict)  # 参数
    ports: List[Dict] = field(default_factory=list)  # 端口信息
    description: str = ""
    

@dataclass
class SignalLine:
    """信号线"""
    src_block: str              # 源模块 SID
    src_port: int               # 源端口
    dst_block: str              # 目标模块 SID
    dst_port: int               # 目标端口
    name: str = ""              # 信号名称
    

@dataclass
class SubSystem:
    """子系统信息"""
    sid: str                    # 系统 ID（实际上是父 Block 的 SID）
    name: str                   # 子系统名称
    parent: str = ""            # 父系统 SID
    blocks: List[Block] = field(default_factory=list)
    inports: List[Port] = field(default_factory=list)
    outports: List[Port] = field(default_factory=list)
    signals: List[SignalLine] = field(default_factory=list)
    description: str = ""
    children: List['SubSystem'] = field(default_factory=list)


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


class SLXParser:
    """SLX 文件解析器"""
    
    def __init__(self, slx_path: str):
        self.slx_path = Path(slx_path)
        self.model = None
        
    def parse(self) -> SimulinkModel:
        """解析 SLX 文件，返回模型对象"""
        if not self.slx_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.slx_path}")
            
        # 检查是否是有效的 ZIP 文件
        if not zipfile.is_zipfile(self.slx_path):
            raise ValueError(f"不是有效的 SLX 文件: {self.slx_path}")
            
        with zipfile.ZipFile(self.slx_path, 'r') as zf:
            # 1. 提取模型基本信息
            self.model = self._parse_metadata(zf)
            
            # 2. 解析模型结构
            self._parse_blockdiagram(zf)
            
        return self.model
    
    def _parse_metadata(self, zf: zipfile.ZipFile) -> SimulinkModel:
        """从 ZIP 中提取模型元信息"""
        model_name = self.slx_path.stem
        version = ""
        created = ""
        modified = ""
        author = ""
        
        # 尝试读取 coreProperties.xml
        try:
            core_xml = zf.read('metadata/coreProperties.xml')
            root = ET.fromstring(core_xml)
            
            # 定义命名空间
            ns = {
                'dc': 'http://purl.org/dc/elements/1.1/',
                'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                'dcterms': 'http://purl.org/dc/terms/'
            }
            
            # 提取创建时间
            created_elem = root.find('.//dcterms:created', ns)
            if created_elem is not None:
                created = created_elem.text or ""
                
            # 提取修改时间
            modified_elem = root.find('.//dcterms:modified', ns)
            if modified_elem is not None:
                modified = modified_elem.text or ""
                
            # 提取作者
            creator_elem = root.find('.//dc:creator', ns)
            if creator_elem is not None:
                author = creator_elem.text or ""
                
        except KeyError:
            pass  # 文件可能不存在，使用默认值
            
        return SimulinkModel(
            name=model_name,
            version=version,
            created=created,
            modified=modified,
            author=author
        )
    
    def _parse_blockdiagram(self, zf: zipfile.ZipFile):
        """解析 blockdiagram.xml 文件"""
        # 尝试读取 blockdiagram.xml
        try:
            xml_content = zf.read('simulink/blockdiagram.xml')
        except KeyError:
            raise ValueError("无法找到 simulink/blockdiagram.xml 文件")
            
        # 解析 XML
        root = ET.fromstring(xml_content)
        
        # 查找 Model 节点
        # <ModelInformation><Model>
        model_elem = root.find('.//Model')
        if model_elem is None:
            raise ValueError("无法找到 Model 节点")
            
        # 获取模型版本
        version_elem = model_elem.find('.//P[@Name="ComputedModelVersion"]')
        if version_elem is not None:
            self.model.version = version_elem.text or ""
        
        # 查找根 System 节点
        # System 节点在 Model 下面，包含所有顶层 Block
        system_elem = model_elem.find('.//System')
        if system_elem is None:
            raise ValueError("无法找到根 System 节点")
            
        # 解析根系统
        self.model.root_system = self._parse_system(system_elem, parent_sid="", system_name="Root")
        
        # 构建子系统字典
        self._build_subsystem_dict(self.model.root_system)
        
    def _parse_system(self, elem: ET.Element, parent_sid: str = "", system_name: str = "Root") -> SubSystem:
        """
        解析 System 节点
        
        注意：System 节点本身没有 SID，SID 是在包含它的 Block 上
        这里我们用 parent_sid 来标识（即父 Block 的 SID）
        """
        subsystem = SubSystem(
            sid=parent_sid,
            name=system_name,
            parent=parent_sid
        )
        
        # 遍历所有 Block
        for block_elem in elem.findall('Block'):
            block = self._parse_block(block_elem)
            subsystem.blocks.append(block)
            
            # 检查是否是 Inport/Outport
            if block.block_type == 'Inport':
                port = self._create_port_from_block(block, 'inport')
                subsystem.inports.append(port)
            elif block.block_type == 'Outport':
                port = self._create_port_from_block(block, 'outport')
                subsystem.outports.append(port)
                
            # 检查是否是 SubSystem，如果是，递归解析内部 System
            if block.block_type == 'SubSystem':
                inner_system_elem = block_elem.find('System')
                if inner_system_elem is not None:
                    child_subsystem = self._parse_system(
                        inner_system_elem, 
                        parent_sid=block.sid,
                        system_name=block.name
                    )
                    subsystem.children.append(child_subsystem)
        
        return subsystem
    
    def _parse_block(self, elem: ET.Element) -> Block:
        """解析 Block 节点"""
        sid = elem.get('SID', '')
        name = elem.get('Name', '')
        block_type = elem.get('BlockType', 'Unknown')
        
        # 解析参数
        parameters = {}
        for p_elem in elem.findall('P'):
            param_name = p_elem.get('Name', '')
            param_value = p_elem.text or ''
            if param_name:
                parameters[param_name] = param_value
        
        # 解析位置
        position = (0, 0, 0, 0)
        pos_str = parameters.get('Position', '')
        if pos_str:
            try:
                # 格式: [x1, y1, x2, y2]
                coords = eval(pos_str)
                if isinstance(coords, (list, tuple)) and len(coords) >= 4:
                    position = tuple(coords[:4])
            except:
                pass
        
        # 解析端口信息
        ports = []
        for port_elem in elem.findall('Port'):
            port_info = {}
            for p_elem in port_elem.findall('P'):
                port_info[p_elem.get('Name', '')] = p_elem.text or ''
            ports.append(port_info)
        
        # 获取描述
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
        # 端口号
        port_number = 1
        if 'Port' in block.parameters:
            try:
                port_number = int(block.parameters['Port'])
            except:
                pass
        
        # 如果有 Port 节点，优先使用里面的信息
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
        
        # 数据类型
        data_type = block.parameters.get('OutDataTypeStr', '')
        
        return Port(
            name=port_name,
            port_type=port_type,
            port_number=port_number,
            data_type=data_type,
            description=block.description
        )
    
    def _build_subsystem_dict(self, subsystem: SubSystem):
        """构建子系统字典，方便查找"""
        if subsystem.sid:
            self.model.all_subsystems[subsystem.sid] = subsystem
            
        for child in subsystem.children:
            self._build_subsystem_dict(child)


def parse_slx(slx_path: str) -> SimulinkModel:
    """便捷函数：解析 SLX 文件"""
    parser = SLXParser(slx_path)
    return parser.parse()


if __name__ == "__main__":
    # 测试代码
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
            print(f"子系统数: {len(model.root_system.children)}")
            
            # 显示前几个模块
            print("\n顶层模块:")
            for block in model.root_system.blocks[:10]:
                print(f"  - {block.name} ({block.block_type})")
                
            # 显示子系统
            if model.root_system.children:
                print("\n子系统:")
                for child in model.root_system.children:
                    print(f"  - {child.name} ({len(child.blocks)} 个模块)")
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()