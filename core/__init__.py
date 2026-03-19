"""
Simulink 文档生成器核心模块
"""

from .parser import SLXParser, SimulinkModel, SubSystem, Block, Port, SignalLine, parse_slx
from .analyzer import ModelAnalyzer, analyze_model
from .generator import DocGenerator, generate_document

__all__ = [
    'SLXParser',
    'SimulinkModel', 
    'SubSystem',
    'Block',
    'Port',
    'SignalLine',
    'parse_slx',
    'ModelAnalyzer',
    'analyze_model',
    'DocGenerator',
    'generate_document'
]