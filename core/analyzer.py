"""
模型分析器（增强版）
分析 Simulink 模型结构，提取信号流、识别功能块、推断算法
"""

from typing import Dict, List, Set, Tuple, Optional
from dataclasses import dataclass
from collections import defaultdict
import re

from .parser import SimulinkModel, SubSystem, Block, SignalLine, Port


@dataclass
class SignalPath:
    """信号路径"""
    start_block: str
    end_block: str
    path: List[str]
    signal_name: str = ""


@dataclass
class BlockTypeStats:
    """模块类型统计"""
    block_type: str
    count: int
    examples: List[str]


@dataclass
class FunctionBlock:
    """功能块"""
    name: str
    block_type: str
    function: str
    inputs: List[str]
    outputs: List[str]
    parameters: Dict


class ModelAnalyzer:
    """模型分析器"""
    
    def __init__(self, model: SimulinkModel):
        self.model = model
        self._block_dict: Dict[str, Block] = {}
        self._subsystem_dict: Dict[str, SubSystem] = {}
        
        if model.root_system:
            self._build_dicts(model.root_system)
    
    def _build_dicts(self, subsystem: SubSystem):
        """构建查找字典"""
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
                examples=names[:5]
            ))
            
        stats.sort(key=lambda x: x.count, reverse=True)
        return stats
    
    def get_all_subsystems(self) -> List[Tuple[int, SubSystem]]:
        """获取所有子系统列表"""
        if not self.model.root_system:
            return []
            
        subsystems = []
        self._collect_subsystems(self.model.root_system, subsystems, level=0)
        return subsystems
    
    def _collect_subsystems(self, subsystem: SubSystem, result: List[Tuple[int, SubSystem]], level: int):
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
    
    # ==================== 新增：信号流分析 ====================
    
    def analyze_signal_flow(self, subsystem: SubSystem) -> Dict:
        """分析子系统的信号流"""
        result = {
            'input_signals': [],
            'output_signals': [],
            'internal_signals': [],
            'signal_chains': [],
            'function_blocks': []
        }
        
        # 获取子系统内的模块字典
        blocks_by_sid = {b.sid: b for b in subsystem.blocks}
        
        # 分析输入信号流
        for inport in subsystem.inports:
            inport_block = self._find_block_by_name(subsystem, inport.name)
            if inport_block:
                destinations = self._trace_signal_forward(inport_block.sid, subsystem, blocks_by_sid, set())
                result['input_signals'].append({
                    'port': inport.name,
                    'connects_to': destinations
                })
        
        # 分析输出信号流
        for outport in subsystem.outports:
            outport_block = self._find_block_by_name(subsystem, outport.name)
            if outport_block:
                sources = self._trace_signal_backward(outport_block.sid, subsystem, blocks_by_sid, set())
                result['output_signals'].append({
                    'port': outport.name,
                    'connects_from': sources
                })
        
        # 识别功能块
        result['function_blocks'] = self._identify_function_blocks(subsystem)
        
        # 分析信号链
        result['signal_chains'] = self._analyze_signal_chains(subsystem)
        
        return result
    
    def _find_block_by_name(self, subsystem: SubSystem, name: str) -> Optional[Block]:
        """按名称查找模块"""
        for block in subsystem.blocks:
            if block.name == name:
                return block
        return None
    
    def _trace_signal_forward(self, sid: str, subsystem: SubSystem, blocks_by_sid: Dict, visited: Set) -> List[str]:
        """向前追踪信号"""
        if sid in visited:
            return []
        visited.add(sid)
        
        destinations = []
        if sid in subsystem.signal_graph:
            for dst_sid, src_port, dst_port in subsystem.signal_graph[sid]:
                dst_block = blocks_by_sid.get(dst_sid)
                if dst_block:
                    if dst_block.block_type == 'Outport':
                        destinations.append(f"输出[{dst_block.name}]")
                    else:
                        destinations.append(dst_block.name)
                        # 继续追踪
                        next_dests = self._trace_signal_forward(dst_sid, subsystem, blocks_by_sid, visited)
                        destinations.extend(next_dests)
        
        return destinations
    
    def _trace_signal_backward(self, sid: str, subsystem: SubSystem, blocks_by_sid: Dict, visited: Set) -> List[str]:
        """向后追踪信号"""
        if sid in visited:
            return []
        visited.add(sid)
        
        sources = []
        if sid in subsystem.reverse_signal_graph:
            for src_sid, src_port, dst_port in subsystem.reverse_signal_graph[sid]:
                src_block = blocks_by_sid.get(src_sid)
                if src_block:
                    if src_block.block_type == 'Inport':
                        sources.append(f"输入[{src_block.name}]")
                    else:
                        sources.append(src_block.name)
                        # 继续追踪
                        prev_sources = self._trace_signal_backward(src_sid, subsystem, blocks_by_sid, visited)
                        sources.extend(prev_sources)
        
        return sources
    
    def _identify_function_blocks(self, subsystem: SubSystem) -> List[Dict]:
        """识别功能块"""
        function_blocks = []
        blocks_by_sid = {b.sid: b for b in subsystem.blocks}
        
        # 已处理的模块
        processed = set()
        
        # 识别 PI 控制器模式
        pi_controllers = self._find_pi_controllers(subsystem, blocks_by_sid, processed)
        function_blocks.extend(pi_controllers)
        
        # 识别滤波器模式
        filters = self._find_filters(subsystem, blocks_by_sid, processed)
        function_blocks.extend(filters)
        
        # 识别坐标变换模式
        transforms = self._find_transforms(subsystem, blocks_by_sid, processed)
        function_blocks.extend(transforms)
        
        # 识别算术运算
        arithmetic = self._find_arithmetic_blocks(subsystem, blocks_by_sid, processed)
        function_blocks.extend(arithmetic)
        
        return function_blocks
    
    def _find_pi_controllers(self, subsystem: SubSystem, blocks_by_sid: Dict, processed: Set) -> List[Dict]:
        """识别 PI 控制器"""
        pi_controllers = []
        
        # PI 控制器通常包含：Sum(误差), Gain(Kp), Gain(Ki), Integrator, Sum(输出)
        # 简化识别：查找 Sum + Gain 组合
        
        sum_blocks = [b for b in subsystem.blocks if b.block_type == 'Sum' and b.sid not in processed]
        
        for sum_block in sum_blocks:
            # 查找连接到该 Sum 的 Gain 模块
            connected_gains = []
            if sum_block.sid in subsystem.signal_graph:
                for dst_sid, _, _ in subsystem.signal_graph[sum_block.sid]:
                    dst_block = blocks_by_sid.get(dst_sid)
                    if dst_block and dst_block.block_type == 'Gain':
                        connected_gains.append(dst_block)
            
            if len(connected_gains) >= 1:
                # 可能是 PI 或 P 控制器
                kp = None
                ki = None
                for gain in connected_gains:
                    gain_value = gain.parameters.get('Gain', '')
                    name_lower = gain.name.lower()
                    if 'kp' in name_lower or 'p' in name_lower:
                        kp = gain_value
                    elif 'ki' in name_lower or 'i' in name_lower:
                        ki = gain_value
                
                if kp or ki:
                    processed.add(sum_block.sid)
                    for g in connected_gains:
                        processed.add(g.sid)
                    
                    pi_controllers.append({
                        'type': 'PI控制器' if ki else 'P控制器',
                        'components': [sum_block.name] + [g.name for g in connected_gains],
                        'parameters': {
                            'Kp': kp or 'N/A',
                            'Ki': ki or 'N/A'
                        },
                        'description': f"{'PI' if ki else 'P'}控制器，误差校正"
                    })
        
        return pi_controllers
    
    def _find_filters(self, subsystem: SubSystem, blocks_by_sid: Dict, processed: Set) -> List[Dict]:
        """识别滤波器"""
        filters = []
        
        # 查找 Delay, UnitDelay, TransferFcn 等模块
        for block in subsystem.blocks:
            if block.sid in processed:
                continue
                
            if block.block_type in ('TransferFcn', 'DiscreteFilter', 'Filter'):
                processed.add(block.sid)
                filters.append({
                    'type': '滤波器',
                    'components': [block.name],
                    'parameters': block.parameters,
                    'description': '信号滤波处理'
                })
            elif block.block_type in ('Delay', 'UnitDelay', 'IntegerDelay'):
                processed.add(block.sid)
                filters.append({
                    'type': '延迟单元',
                    'components': [block.name],
                    'parameters': block.parameters,
                    'description': '信号延迟处理'
                })
        
        return filters
    
    def _find_transforms(self, subsystem: SubSystem, blocks_by_sid: Dict, processed: Set) -> List[Dict]:
        """识别坐标变换"""
        transforms = []
        
        # 根据名称识别
        transform_keywords = ['park', 'clark', 'inverse', 'transform', 'dq', 'alpha', 'beta']
        
        for block in subsystem.blocks:
            if block.sid in processed:
                continue
            
            name_lower = block.name.lower()
            
            if 'park' in name_lower and 'inv' in name_lower:
                processed.add(block.sid)
                transforms.append({
                    'type': '逆Park变换',
                    'components': [block.name],
                    'description': 'dq坐标系 → αβ坐标系'
                })
            elif 'park' in name_lower:
                processed.add(block.sid)
                transforms.append({
                    'type': 'Park变换',
                    'components': [block.name],
                    'description': 'αβ坐标系 → dq坐标系'
                })
            elif 'clark' in name_lower and ('inv' in name_lower or 'inverse' in name_lower):
                processed.add(block.sid)
                transforms.append({
                    'type': '逆Clark变换',
                    'components': [block.name],
                    'description': 'αβ坐标系 → abc坐标系'
                })
            elif 'clark' in name_lower:
                processed.add(block.sid)
                transforms.append({
                    'type': 'Clark变换',
                    'components': [block.name],
                    'description': 'abc坐标系 → αβ坐标系'
                })
            elif 'svpwm' in name_lower:
                processed.add(block.sid)
                transforms.append({
                    'type': 'SVPWM',
                    'components': [block.name],
                    'description': '空间矢量脉宽调制'
                })
        
        return transforms
    
    def _find_arithmetic_blocks(self, subsystem: SubSystem, blocks_by_sid: Dict, processed: Set) -> List[Dict]:
        """识别算术运算模块"""
        arithmetic = []
        
        for block in subsystem.blocks:
            if block.sid in processed:
                continue
            
            if block.block_type == 'Gain':
                gain_value = block.parameters.get('Gain', 'N/A')
                arithmetic.append({
                    'type': '增益',
                    'components': [block.name],
                    'parameters': {'Gain': gain_value},
                    'description': f"增益放大，系数={gain_value}"
                })
            elif block.block_type == 'Sum':
                arithmetic.append({
                    'type': '求和',
                    'components': [block.name],
                    'description': '信号求和运算'
                })
            elif block.block_type == 'Product':
                arithmetic.append({
                    'type': '乘法',
                    'components': [block.name],
                    'description': '信号乘法运算'
                })
            elif block.block_type == 'TrigonometricFunction':
                func = block.parameters.get('Operator', 'sin')
                arithmetic.append({
                    'type': '三角函数',
                    'components': [block.name],
                    'parameters': {'Function': func},
                    'description': f'三角函数运算（{func}）'
                })
        
        return arithmetic
    
    def _analyze_signal_chains(self, subsystem: SubSystem) -> List[Dict]:
        """分析信号链"""
        chains = []
        blocks_by_sid = {b.sid: b for b in subsystem.blocks}
        
        # 从输入端口追踪到输出端口
        for inport in subsystem.inports:
            inport_block = self._find_block_by_name(subsystem, inport.name)
            if not inport_block:
                continue
            
            # 追踪信号路径
            path = self._trace_full_path(inport_block.sid, subsystem, blocks_by_sid, [])
            
            if path:
                chains.append({
                    'input': inport.name,
                    'path': path,
                    'output': path[-1] if path else ''
                })
        
        return chains
    
    def _trace_full_path(self, sid: str, subsystem: SubSystem, blocks_by_sid: Dict, path: List[str], visited: Set = None) -> List[str]:
        """追踪完整信号路径"""
        if visited is None:
            visited = set()
        
        if sid in visited:
            return path
        
        visited.add(sid)
        block = blocks_by_sid.get(sid)
        
        if block:
            path.append(block.name)
        
        # 继续追踪
        if sid in subsystem.signal_graph:
            for dst_sid, _, _ in subsystem.signal_graph[sid]:
                dst_block = blocks_by_sid.get(dst_sid)
                if dst_block and dst_block.block_type not in ('Outport',):
                    path = self._trace_full_path(dst_sid, subsystem, blocks_by_sid, path, visited)
                elif dst_block and dst_block.block_type == 'Outport':
                    path.append(f"→ 输出[{dst_block.name}]")
        
        return path
    
    def get_detailed_function_description(self, subsystem: SubSystem) -> str:
        """生成详细的功能描述"""
        parts = []
        
        # 1. 基本信息
        parts.append(f"### 子系统：{subsystem.name}")
        parts.append(f"包含 {len(subsystem.blocks)} 个模块，{len(subsystem.inports)} 个输入端口，{len(subsystem.outports)} 个输出端口。")
        
        # 2. 分析信号流
        signal_flow = self.analyze_signal_flow(subsystem)
        
        # 3. 功能块描述
        if signal_flow['function_blocks']:
            parts.append("\n### 功能模块")
            for fb in signal_flow['function_blocks']:
                parts.append(f"- **{fb['type']}**：{', '.join(fb['components'])}")
                if fb.get('description'):
                    parts.append(f"  - {fb['description']}")
                if fb.get('parameters'):
                    params_str = ', '.join(f"{k}={v}" for k, v in fb['parameters'].items())
                    parts.append(f"  - 参数：{params_str}")
        
        # 4. 信号链描述
        if signal_flow['signal_chains']:
            parts.append("\n### 信号处理流程")
            for chain in signal_flow['signal_chains']:
                path_str = ' → '.join(chain['path'])
                parts.append(f"- {chain['input']}: {path_str}")
        
        # 5. 子系统列表
        if subsystem.children:
            parts.append("\n### 包含的子系统")
            for child in subsystem.children:
                parts.append(f"- {child.name} ({len(child.blocks)} 模块)")
        
        return '\n'.join(parts)


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