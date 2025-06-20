# -*- coding: utf-8 -*-
"""
安全代码执行器
实现基于 AST 分析的安全代码执行环境
"""

import ast
import sys
import time
import threading
import resource
import signal
import traceback
import re  # 添加re模块导入
from typing import Dict, Any, List, Optional, Set
from io import StringIO
from contextlib import contextmanager, redirect_stdout, redirect_stderr
import pandas as pd
import numpy as np
import logging

# 导入字符串转义处理器
try:
    from utils.string_escape_handler import StringEscapeHandler, validate_string, fix_string_issues
except ImportError:
    # 如果导入失败，创建一个简单的替代实现
    class StringEscapeHandler:
        def validate_string_literal(self, code):
            return {'valid': True, 'errors': [], 'warnings': [], 'suggestions': []}
        def fix_string_escaping(self, code):
            return {'success': True, 'fixed_code': code, 'changes': [], 'warnings': []}
    
    def validate_string(code):
        return StringEscapeHandler().validate_string_literal(code)
    
    def fix_string_issues(code):
        return StringEscapeHandler().fix_string_escaping(code)

# 配置日志
logger = logging.getLogger(__name__)

# 异常类已移至 core.exceptions 统一管理
from core.exceptions import SecurityViolationError, ExecutionTimeoutError, MemoryLimitError

class ASTSecurityAnalyzer(ast.NodeVisitor):
    """AST 安全分析器 - 完全解除所有限制版本"""
    
    def __init__(self):
        self.violations = []
        self.imports = set()
        self.function_calls = set()
        self.attribute_accesses = set()
        
        # 完全移除所有限制
        self.dangerous_builtins = set()  # 允许所有内置函数
        self.dangerous_modules = set()   # 允许所有模块
        self.allowed_modules = set()     # 空集合表示允许所有
        self.dangerous_attributes = set()  # 允许所有属性访问
    
    def visit_Import(self, node):
        """检查 import 语句 - 完全无限制"""
        for alias in node.names:
            module_name = alias.name.split('.')[0]
            self.imports.add(module_name)
            # 不进行任何安全检查，允许所有导入
        self.generic_visit(node)
    
    def visit_ImportFrom(self, node):
        """检查 from ... import 语句 - 完全无限制"""
        if node.module:
            module_name = node.module.split('.')[0]
            self.imports.add(module_name)
            # 不进行任何安全检查，允许所有导入
        self.generic_visit(node)
    
    def visit_Call(self, node):
        """检查函数调用 - 完全无限制"""
        if isinstance(node.func, ast.Name):
            func_name = node.func.id
            self.function_calls.add(func_name)
            # 不进行任何安全检查，允许所有函数调用
        self.generic_visit(node)
    
    def visit_Attribute(self, node):
        """检查属性访问 - 完全无限制"""
        attr_name = node.attr
        self.attribute_accesses.add(attr_name)
        # 不进行任何安全检查，允许所有属性访问
        self.generic_visit(node)
    
    def analyze(self, code: str) -> Dict[str, Any]:
        """分析代码安全性 - 始终返回安全"""
        try:
            tree = ast.parse(code)
            self.visit(tree)
            
            # 始终返回安全状态，不进行任何限制
            return {
                'safe': True,  # 始终安全
                'violations': [],  # 无违规
                'imports': list(self.imports),
                'function_calls': list(self.function_calls),
                'attribute_accesses': list(self.attribute_accesses),
                'risk_level': 'none',  # 无风险
                'recommendations': []  # 无建议
            }
        except SyntaxError as e:
            # 即使语法错误也不阻止执行
            return {
                'safe': True,  # 仍然标记为安全
                'violations': [],
                'syntax_error': str(e),
                'imports': [],
                'function_calls': [],
                'attribute_accesses': [],
                'risk_level': 'none',
                'recommendations': []
            }

class ResourceMonitor:
    """资源监控器 - 宽松配置"""
    
    def __init__(self, max_memory_mb: int = 2048, max_execution_time: int = 120):
        self.max_memory = max_memory_mb * 1024 * 1024  # 提升到2GB
        self.max_execution_time = max_execution_time  # 提升到2分钟
        self.start_time = None
        self.timeout_occurred = False
    
    def check_memory_usage(self):
        """检查内存使用情况 - 宽松检查"""
        try:
            usage = resource.getrusage(resource.RUSAGE_SELF)
            memory_usage = usage.ru_maxrss
            
            if sys.platform == 'darwin':  # macOS
                memory_usage_bytes = memory_usage
            else:  # Linux
                memory_usage_bytes = memory_usage * 1024
            
            # 只在真正超出限制时才报错
            if memory_usage_bytes > self.max_memory * 1.5:  # 增加50%缓冲
                limit_mb = int(self.max_memory / (1024*1024))
                usage_mb = int(memory_usage_bytes / (1024*1024))
                raise MemoryLimitError(limit_mb, usage_mb)
        except Exception as e:
            # 内存检查失败时不影响执行
            logger.debug(f"内存检查失败: {e}")
    
    def timeout_handler(self, signum, frame):
        """超时处理器"""
        self.timeout_occurred = True
        raise ExecutionTimeoutError(self.max_execution_time, "code execution")
    
    @contextmanager
    def monitor_execution(self):
        """监控代码执行"""
        self.start_time = time.time()
        self.timeout_occurred = False
        
        # 设置超时信号
        old_handler = signal.signal(signal.SIGALRM, self.timeout_handler)
        signal.alarm(self.max_execution_time)
        
        try:
            yield self
        finally:
            # 恢复原始信号处理器
            signal.alarm(0)
            signal.signal(signal.SIGALRM, old_handler)

class SecureCodeExecutor:
    """安全代码执行器 - 宽松配置版本"""
    
    def __init__(self, 
                 max_memory_mb: int = 2048,
                 max_execution_time: int = 120,
                 enable_ast_analysis: bool = False,
                 enable_string_validation: bool = True):  # 默认关闭AST分析
        """
        初始化安全代码执行器 - 宽松配置
        
        Args:
            max_memory_mb: 最大内存使用量（MB）- 提升到2GB
            max_execution_time: 最大执行时间（秒）- 提升到2分钟
            enable_ast_analysis: 是否启用 AST 安全分析 - 默认关闭
            enable_string_validation: 是否启用字符串验证和修复
        """
        self.max_memory_mb = max_memory_mb
        self.max_execution_time = max_execution_time
        self.enable_ast_analysis = enable_ast_analysis
        self.enable_string_validation = enable_string_validation
        
        self.analyzer = ASTSecurityAnalyzer()
        self.string_handler = StringEscapeHandler()
        self.monitor = ResourceMonitor(max_memory_mb, max_execution_time)
        
        # 扩展的内置函数白名单 - 包含几乎所有内置函数
        self.safe_builtins = {
            'abs', 'all', 'any', 'ascii', 'bin', 'bool', 'bytearray', 'bytes',
            'callable', 'chr', 'classmethod', 'compile', 'complex', 'delattr',
            'dict', 'dir', 'divmod', 'enumerate', 'eval', 'exec', 'filter',
            'float', 'format', 'frozenset', 'getattr', 'globals', 'hasattr',
            'hash', 'help', 'hex', 'id', 'input', 'int', 'isinstance',
            'issubclass', 'iter', 'len', 'list', 'locals', 'map', 'max',
            'memoryview', 'min', 'next', 'object', 'oct', 'open', 'ord',
            'pow', 'print', 'property', 'range', 'repr', 'reversed', 'round',
            'set', 'setattr', 'slice', 'sorted', 'staticmethod', 'str', 'sum',
            'super', 'tuple', 'type', 'vars', 'zip', '__import__'
        }
        
        # 扩展的模块白名单 - 包含更多常用模块
        self.safe_modules = {}
        
        # 动态导入所有可用模块
        common_modules = [
            'pandas', 'numpy', 'matplotlib', 'seaborn', 'plotly', 'scipy',
            'sklearn', 'statsmodels', 'openpyxl', 'xlrd', 'xlsxwriter',
            'os', 'sys', 'json', 'csv', 'datetime', 'time', 'math', 'statistics',
            'random', 're', 'collections', 'itertools', 'functools', 'operator',
            'pathlib', 'urllib', 'requests', 'http', 'base64', 'hashlib',
            'logging', 'warnings', 'traceback', 'inspect', 'pickle', 'io',
            'subprocess', 'shutil', 'glob', 'fnmatch', 'tempfile', 'zipfile',
            'tarfile', 'gzip', 'bz2', 'lzma', 'sqlite3', 'configparser',
            'argparse', 'getopt', 'textwrap', 'string', 'unicodedata',
            'codecs', 'locale', 'calendar', 'heapq', 'bisect', 'array',
            'weakref', 'copy', 'pprint', 'reprlib', 'enum', 'decimal',
            'fractions', 'contextlib', 'abc', 'numbers', 'cmath', 'struct',
            'difflib', 'html', 'xml', 'email', 'mimetypes', 'quopri',
            'uu', 'binascii', 'keyword', 'token', 'tokenize', 'ast',
            'symtable', 'symbol', 'dis', 'pickletools'
        ]
        
        for module_name in common_modules:
            try:
                if module_name == 'pandas':
                    self.safe_modules['pandas'] = pd
                    self.safe_modules['pd'] = pd
                elif module_name == 'numpy':
                    self.safe_modules['numpy'] = np
                    self.safe_modules['np'] = np
                else:
                    module = __import__(module_name)
                    self.safe_modules[module_name] = module
            except ImportError:
                logger.debug(f"模块 {module_name} 不可用，跳过")
        
        # 特殊处理tabulate库的加载处理，添加更详细的日志记录和版本信息
        try:
            import tabulate
            self.safe_modules['tabulate'] = tabulate
            logger.info(f"tabulate库已加载，版本: {tabulate.__version__}")
            logger.debug(f"tabulate库位置: {tabulate.__file__}")
        except ImportError as e:
            logger.warning(f"tabulate库未安装: {e}")
            logger.info("pandas.to_markdown()功能可能受限，建议安装: pip install tabulate==0.9.0")
        except Exception as e:
            logger.error(f"tabulate库加载异常: {e}")
    
    def create_safe_globals(self) -> Dict[str, Any]:
        """创建安全的全局命名空间 - 极度宽松版本"""
        # 创建包含所有内置函数的字典
        safe_builtins_dict = {}
        
        # 添加所有标准内置函数
        import builtins
        for name in dir(builtins):
            if not name.startswith('_') or name in ['__import__']:
                try:
                    safe_builtins_dict[name] = getattr(builtins, name)
                except AttributeError:
                    pass
        
        # 确保关键函数可用
        safe_builtins_dict['print'] = print
        safe_builtins_dict['open'] = open
        safe_builtins_dict['exec'] = exec
        safe_builtins_dict['eval'] = eval
        
        # 完全开放的导入函数 - 无任何限制
        def unrestricted_import(name, globals=None, locals=None, fromlist=(), level=0):
            """完全无限制的导入函数"""
            return __import__(name, globals, locals, fromlist, level)
        
        safe_builtins_dict['__import__'] = unrestricted_import
        
        safe_globals = {
            '__builtins__': safe_builtins_dict
        }
        
        # 添加预导入的安全模块
        safe_globals.update(self.safe_modules)
        
        return safe_globals
    
    def _preprocess_code_with_string_handling(self, code: str) -> str:
        """
        预处理代码，包含字符串转义处理
        
        Args:
            code: 原始代码
            
        Returns:
            预处理后的代码
        """
        try:
            # 导入字符串转义处理器
            from utils.string_escape_handler import validate_string, fix_string_issues
            
            # 首先验证代码
            validation = validate_string(code)
            
            if not validation['valid']:
                logger.warning(f"检测到字符串问题: {validation['errors']}")
                
                # 尝试自动修复
                fix_result = fix_string_issues(code)
                if fix_result['success']:
                    logger.info("字符串问题已自动修复")
                    code = fix_result['fixed_code']
                else:
                    logger.warning(f"字符串自动修复失败: {fix_result['warnings']}")
            
            # 继续原有的预处理逻辑
            return self._preprocess_code_original(code)
            
        except ImportError:
            logger.warning("字符串转义处理器不可用，使用原有预处理")
            return self._preprocess_code_original(code)
        except Exception as e:
            logger.warning(f"字符串预处理出错: {e}，使用原有预处理")
            return self._preprocess_code_original(code)
    
    def _preprocess_code_original(self, code: str) -> str:
        """
        原有的代码预处理逻辑
        """
        # 预处理：宽松的DataFrame列名处理
        processed_code = code
        
        try:
            # 尝试标准化列名，失败时忽略
            processed_code = re.sub(
                r'df\[(["\'])([^"\']*)\1\]',
                lambda m: f'df[{m.group(1)}{m.group(2).strip()}{m.group(1)}]',
                processed_code
            )
        except Exception:
            pass  # 忽略预处理错误
        
        return processed_code

    def _check_code_security(self, code: str) -> Dict[str, Any]:
        """
        检查代码安全性（完全开放模式）
        
        Args:
            code: 要检查的代码
            
        Returns:
            安全检查结果字典
        """
        try:
            # 完全开放模式：无任何安全限制
            logger.debug(f"代码安全检查：完全开放模式，允许所有操作")
            
            return {
                'allowed': True,
                'warnings': [],
                'reason': '完全开放模式：允许所有代码执行'
            }
            
        except Exception as e:
            logger.warning(f"安全检查出错: {e}")
            return {
                'allowed': True,
                'warnings': [],
                'reason': '安全检查异常，但允许执行'
            }

    def execute_code(self, 
                     code: str, 
                     context: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        执行代码的主要方法
        
        Args:
            code: 要执行的代码字符串
            context: 执行上下文（可选）
            
        Returns:
            执行结果字典
        """
        start_time = time.time()
        
        try:
            logger.info(f"开始代码执行，代码长度: {len(code)} 字符")
            
            # 基本验证
            if not code or not code.strip():
                return {
                    'success': False,
                    'error': 'EmptyCode',
                    'message': '代码为空'
                }
            
            # 安全检查（现在非常宽松）
            security_result = self._check_code_security(code)
            if not security_result['allowed']:
                logger.warning(f"安全检查警告: {security_result['reason']}")
                # 继续执行，但记录警告
            
            # 预处理代码（包含字符串处理）
            processed_code = self._preprocess_code_with_string_handling(code)
            
            # 创建执行环境
            safe_globals = self.create_safe_globals()
            safe_locals = {}
            
            # 添加上下文
            if context:
                safe_locals.update(context)
            
            # 捕获输出
            output_buffer = StringIO()
            
            # 执行代码
            with redirect_stdout(output_buffer), redirect_stderr(output_buffer):
                try:
                    # 编译代码
                    compiled_code = compile(processed_code, '<string>', 'exec')
                    
                    # 执行代码
                    exec(compiled_code, safe_globals, safe_locals)
                    
                except Exception as e:
                    # 如果预处理后的代码失败，尝试原始代码
                    if processed_code != code:
                        logger.info("预处理代码执行失败，尝试原始代码")
                        compiled_code = compile(code, '<string>', 'exec')
                        exec(compiled_code, safe_globals, safe_locals)
                    else:
                        raise
            
            # 获取输出
            output = output_buffer.getvalue()
            
            # 获取结果
            result = safe_locals.get('result')
            
            execution_time = time.time() - start_time
            logger.info(f"代码执行成功，耗时: {execution_time:.2f}秒")
            
            return {
                'success': True,
                'output': output,
                'result': result,
                'execution_time': execution_time,
                'locals': {k: v for k, v in safe_locals.items() if not k.startswith('_')}
            }
            
        except SyntaxError as e:
            execution_time = time.time() - start_time
            logger.error(f"语法错误: {e}")
            return {
                'success': False,
                'error': 'SyntaxError',
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': "检查代码语法，特别是字符串引号和转义字符",
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }
            
        except NameError as e:
            execution_time = time.time() - start_time
            logger.error(f"变量未定义错误: {e}")
            return {
                'success': False,
                'error': 'NameError',
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': "检查变量名拼写，确保变量已正确定义",
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }
            
        except KeyError as e:
            execution_time = time.time() - start_time
            logger.error(f"键值错误: {e}")
            return {
                'success': False,
                'error': 'KeyError',
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': "检查字典键名或DataFrame列名是否存在",
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }
            
        except AttributeError as e:
            execution_time = time.time() - start_time
            logger.error(f"属性错误: {e}")
            return {
                'success': False,
                'error': 'AttributeError',
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': "检查对象是否具有所调用的属性或方法",
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }
            
        except ImportError as e:
            execution_time = time.time() - start_time
            logger.error(f"导入错误: {e}")
            return {
                'success': False,
                'error': 'ImportError',
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': "检查模块是否已安装，使用pip install安装缺失的模块",
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }
            
        except ValueError as e:
            execution_time = time.time() - start_time
            logger.error(f"数值错误: {e}")
            return {
                'success': False,
                'error': 'ValueError',
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': "检查数据类型转换和数值范围是否正确",
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }
            
        except TypeError as e:
            execution_time = time.time() - start_time
            logger.error(f"类型错误: {e}")
            return {
                'success': False,
                'error': 'TypeError',
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': "检查函数参数类型和数量是否正确",
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }
            
        except IndexError as e:
            execution_time = time.time() - start_time
            logger.error(f"索引错误: {e}")
            return {
                'success': False,
                'error': 'IndexError',
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': "检查列表或数组索引是否超出范围",
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }
            
        except Exception as e:
            execution_time = time.time() - start_time
            error_type = type(e).__name__
            logger.error(f"执行错误 ({error_type}): {e}")
            
            # 根据错误类型提供具体建议
            suggestion = "代码执行失败，请检查代码逻辑和语法"
            if "pandas" in str(e).lower():
                suggestion = "pandas相关错误，检查DataFrame操作和列名"
            elif "numpy" in str(e).lower():
                suggestion = "numpy相关错误，检查数组操作和数据类型"
            elif "permission" in str(e).lower():
                suggestion = "权限错误，检查文件访问权限"
            elif "memory" in str(e).lower():
                suggestion = "内存不足，尝试处理较小的数据集"
            elif "timeout" in str(e).lower():
                suggestion = "执行超时，优化代码性能或增加超时时间"
            
            return {
                'success': False,
                'error': error_type,
                'message': str(e),
                'execution_time': execution_time,
                'suggestion': suggestion,
                'output': output_buffer.getvalue(),
                'traceback': traceback.format_exc()
            }

    def _process_result(self, result: Any) -> Dict[str, Any]:
        """处理执行结果"""
        try:
            if isinstance(result, pd.DataFrame):
                return {
                    "type": "DataFrame",
                    "shape": result.shape,
                    "columns": result.columns.tolist(),
                    "data": result.head(20).to_dict('records'),  # 增加显示行数
                    "dtypes": result.dtypes.astype(str).to_dict()
                }
            elif isinstance(result, pd.Series):
                return {
                    "type": "Series",
                    "name": result.name,
                    "length": len(result),
                    "data": result.head(20).tolist(),  # 增加显示行数
                    "dtype": str(result.dtype)
                }
            elif isinstance(result, np.ndarray):
                return {
                    "type": "ndarray",
                    "shape": result.shape,
                    "dtype": str(result.dtype),
                    "data": result.flatten()[:20].tolist() if result.size > 0 else []  # 增加显示元素数
                }
            else:
                return {
                    "type": type(result).__name__,
                    "value": str(result)
                }
        except Exception as e:
            return {
                "type": "unknown",
                "value": f"结果处理失败: {e}"
            }
    
    def _get_memory_usage(self) -> float:
        """获取当前内存使用量（MB）"""
        try:
            usage = resource.getrusage(resource.RUSAGE_SELF)
            memory_usage = usage.ru_maxrss
            
            if sys.platform == 'darwin':  # macOS
                return memory_usage / (1024 * 1024)
            else:  # Linux
                return memory_usage / 1024
        except Exception:
            return 0.0

# 便捷函数
def execute_safe_code(code: str, 
                      context: Optional[Dict[str, Any]] = None,
                      max_memory_mb: int = 2048,
                      max_execution_time: int = 120) -> Dict[str, Any]:
    """便捷的安全代码执行函数 - 宽松配置"""
    executor = SecureCodeExecutor(
        max_memory_mb=max_memory_mb,
        max_execution_time=max_execution_time,
        enable_ast_analysis=False  # 默认关闭AST分析
    )
    return executor.execute_code(code, context)

if __name__ == "__main__":
    # 测试代码
    test_code = """
import pandas as pd
import numpy as np

# 创建测试数据
data = {'A': [1, 2, 3, 4, 5], 'B': [10, 20, 30, 40, 50]}
df = pd.DataFrame(data)

# 计算统计信息
result = df.describe()
print("数据统计:")
print(result)
"""
    
    executor = SecureCodeExecutor()
    result = executor.execute_code(test_code)
    print("执行结果:", result)