"""
模型分析器
分析 Simulink 模型结构，提取信号流、层级关系等信息
"""

from typing import Dict, List, Set, Tuple
from dataclasses import dataclass
from .parser import SimulinkModel, SubSystem, Block, SignalLine, Port


@dataclass
class SignalPath:
    """信号路径"""
    start_port: Port
    end_port: Port
    path: List[str]  # 经过的模块名称
    signal_name: str = ""


@dataclass
class BlockTypeStats:
    """模块类型统计"""
    block_type: str
    count: int
    examples: List[str]  # 示例模块名称


class ModelAnalyzer:
    """模型分析器"""
    
    def __init__(self, model: SimulinkModel):
        self.model = model
        self._block_dict: Dict[str, Block] = {}
        self._subsystem_dict: Dict[str, SubSystem] = {}
        
        # 构建查找字典
        if model.root_system:
            self._build_dicts(model.root_system)
    
    def _build_dicts(self, subsystem: SubSystem):
        """构建模块和子系统的查找字典"""
        self._subsystem_dict[subsystem.sid or subsystem.name] = subsystem
        
        for block in subsystem.blocks:
            self._block_dict[block.sid or block.name] = block
            
        for child in subsystem.children:
            self._build_dicts(child)
    
    def get_hierarchy_tree(self) -> Dict:
        """获取层级树结构"""
        if not self.model.root_system:
            return {}
            
        return self._build_hierarchy(self.model.root_system)
    
    def _build_hierarchy(self, subsystem: SubSystem) -> Dict:
        """递归构建层级结构"""
        node = {
            'name': subsystem.name,
            'sid': subsystem.sid,
            'block_count': len(subsystem.blocks),
            'inport_count': len(subsystem.inports),
            'outport_count': len(subsystem.outports),
            'children': []
        }
        
        for child in subsystem.children:
            node['children'].append(self._build_hierarchy(child))
            
        return node
    
    def get_block_type_statistics(self) -> List[BlockTypeStats]:
        """统计各类型模块数量"""
        type_counts: Dict[str, List[str]] = {}
        
        for block in self._block_dict.values():
            if block.block_type not in type_counts:
                type_counts[block.block_type] = []
            type_counts[block.block_type].append(block.name)
        
        stats = []
        for block_type, names in type_counts.items():
            stats.append(BlockTypeStats(
                block_type=block_type,
                count=len(names),
                examples=names[:5]  # 只保留前5个示例
            ))
            
        # 按数量排序
        stats.sort(key=lambda x: x.count, reverse=True)
        return stats
    
    def analyze_signals(self, subsystem: SubSystem = None) -> List[SignalPath]:
        """分析信号流"""
        if subsystem is None:
            subsystem = self.model.root_system
            
        if not subsystem:
            return []
            
        paths = []
        
        # 从每个 Inport 开始追踪
        for inport in subsystem.inports:
            signal_paths = self._trace_signal_from_port(inport, subsystem, [])
            paths.extend(signal_paths)
            
        return paths
    
    def _trace_signal_from_port(self, start_port: Port, subsystem: SubSystem, 
                                  visited: List[str]) -> List[SignalPath]:
        """从端口追踪信号"""
        paths = []
        
        # 找到连接到该端口的信号线
        for signal in subsystem.signals:
            if signal.dst_block == start_port.name:
                # 找到目标模块
                dst_block = self._block_dict.get(signal.dst_block)
                if dst_block and dst_block.sid not in visited:
                    # 继续追踪
                    visited.append(dst_block.sid)
                    
                    # 查找该模块的输出信号
                    for out_signal in subsystem.signals:
                        if out_signal.src_block == dst_block.name:
                            # 检查是否连接到 Outport
                            if out_signal.dst_block in [p.name for p in subsystem.outports]:
                                end_port = next(
                                    (p for p in subsystem.outports if p.name == out_signal.dst_block),
                                    None
                                )
                                if end_port:
                                    paths.append(SignalPath(
                                        start_port=start_port,
                                        end_port=end_port,
                                        path=visited + [dst_block.name],
                                        signal_name=signal.name
                                    ))
                                    
        return paths
    
    def get_all_subsystems(self) -> List[SubSystem]:
        """获取所有子系统列表（扁平化）"""
        if not self.model.root_system:
            return []
            
        subsystems = []
        self._collect_subsystems(self.model.root_system, subsystems, level=0)
        return subsystems
    
    def _collect_subsystems(self, subsystem: SubSystem, result: List[Tuple[int, SubSystem]], 
                            level: int):
        """收集所有子系统"""
        result.append((level, subsystem))
        for child in subsystem.children:
            self._collect_subsystems(child, result, level + 1)
    
    def get_subsystem_summary(self, subsystem: SubSystem) -> Dict:
        """获取子系统摘要"""
        return {
            'name': subsystem.name,
            'sid': subsystem.sid,
            'description': subsystem.description,
            'block_count': len(subsystem.blocks),
            'inports': [
                {
                    'name': p.name,
                    'number': p.port_number,
                    'data_type': p.data_type,
                    'description': p.description
                }
                for p in subsystem.inports
            ],
            'outports': [
                {
                    'name': p.name,
                    'number': p.port_number,
                    'data_type': p.data_type,
                    'description': p.description
                }
                for p in subsystem.outports
            ],
            'block_types': self._get_block_types_in_subsystem(subsystem),
            'signal_count': len(subsystem.signals)
        }
    
    def _get_block_types_in_subsystem(self, subsystem: SubSystem) -> Dict[str, int]:
        """统计子系统内的模块类型"""
        types = {}
        for block in subsystem.blocks:
            types[block.block_type] = types.get(block.block_type, 0) + 1
        return types
    
    def generate_signal_flow_description(self, subsystem: SubSystem) -> str:
        """生成信号流描述文字"""
        if not subsystem.signals:
            return "该子系统没有信号连接。"
            
        descriptions = []
        
        for signal in subsystem.signals:
            src_block = self._block_dict.get(signal.src_block)
            dst_block = self._block_dict.get(signal.dst_block)
            
            src_name = src_block.name if src_block else signal.src_block
            dst_name = dst_block.name if dst_block else signal.dst_block
            
            if signal.name:
                desc = f"信号 [{signal.name}]: {src_name} (端口{signal.src_port}) → {dst_name} (端口{signal.dst_port})"
            else:
                desc = f"{src_name} (端口{signal.src_port}) → {dst_name} (端口{signal.dst_port})"
                
            descriptions.append(desc)
            
        return "\n".join(descriptions)
    
    def get_model_overview(self) -> Dict:
        """获取模型概览"""
        subsystems = self.get_all_subsystems()
        
        return {
            'name': self.model.name,
            'version': self.model.version,
            'created': self.model.created,
            'modified': self.model.modified,
            'author': self.model.author,
            'description': self.model.description,
            'total_blocks': len(self._block_dict),
            'total_subsystems': len(subsystems),
            'max_depth': max((level for level, _ in subsystems), default=0),
            'block_type_stats': self.get_block_type_statistics()
        }


def analyze_model(model: SimulinkModel) -> Dict:
    """便捷函数：分析模型"""
    analyzer = ModelAnalyzer(model)
    return {
        'overview': analyzer.get_model_overview(),
        'hierarchy': analyzer.get_hierarchy_tree(),
        'subsystems': [
            analyzer.get_subsystem_summary(sub) 
            for level, sub in analyzer.get_all_subsystems()
        ]
    }


if __name__ == "__main__":
    # 测试代码
    print("模型分析器模块")