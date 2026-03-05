"""
扫描功能测试文件
用于测试 NAPS2 扫描功能
"""
import os
import json
import subprocess
from datetime import datetime

# 配置
DEFAULT_NAPS2_PATH = r"D:\Download\NAPS2\App\NAPS2.Console.exe"
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'naps2_config.json')
SCAN_OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'scan')
SCAN_OUTPUT_PATH = os.path.join(SCAN_OUTPUT_DIR, 'scan.png')

def get_naps2_path():
    """获取NAPS2控制台程序路径配置"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                naps2_path = config.get('naps2_path', DEFAULT_NAPS2_PATH)
                return naps2_path.strip()
    except Exception:
        pass
    return DEFAULT_NAPS2_PATH

def initialize_naps2_powershell_function():
    """初始化NAPS2 PowerShell函数"""
    try:
        naps2_path = get_naps2_path()
        
        # 构建PowerShell函数定义脚本
        ps_function_script = f'''
function naps2.console {{
    # 定义NAPS2.Console.exe的完整路径（适配含空格的路径）
    $naps2ConsolePath = "{naps2_path}"
    
    # 检查文件是否存在，避免调用失败
    if (-not (Test-Path -Path $naps2ConsolePath -PathType Leaf)) {{
        Write-Error "错误：未找到NAPS2控制台程序，请检查路径是否正确！路径：$naps2ConsolePath"
        return
    }}
    # 调用程序并传递参数
    & $naps2ConsolePath @args
}}
'''
        
        # 将函数定义写入临时PowerShell脚本文件
        ps_init_file = os.path.join(os.path.dirname(__file__), 'naps2_init.ps1')
        with open(ps_init_file, 'w', encoding='utf-8') as f:
            f.write(ps_function_script)
        
        # 执行PowerShell脚本以定义函数
        result = subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps_init_file],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode != 0:
            print(f"警告：初始化NAPS2 PowerShell函数失败: {result.stderr}")
            return False
        
        print("NAPS2 PowerShell函数初始化成功")
        return True
    except Exception as e:
        print(f"初始化NAPS2 PowerShell函数时出错: {e}")
        return False

def test_scan():
    """测试扫描功能"""
    print("=" * 60)
    print("NAPS2 扫描功能测试")
    print("=" * 60)
    
    # 1. 获取配置路径
    print(f"\n[1] 获取NAPS2路径配置...")
    naps2_path = get_naps2_path()
    print(f"    当前配置路径: {naps2_path}")
    
    # 2. 检查程序是否存在
    print(f"\n[2] 检查NAPS2程序是否存在...")
    if os.path.exists(naps2_path):
        print(f"    ✓ NAPS2程序存在")
    else:
        print(f"    ✗ NAPS2程序不存在！请检查路径是否正确。")
        print(f"    提示：请在设置中配置正确的路径，或修改此文件中的 DEFAULT_NAPS2_PATH")
        return False
    
    # 3. 初始化PowerShell函数
    print(f"\n[3] 初始化PowerShell函数...")
    if initialize_naps2_powershell_function():
        print(f"    ✓ PowerShell函数初始化成功")
    else:
        print(f"    ✗ PowerShell函数初始化失败")
        return False
    
    # 4. 创建输出目录
    print(f"\n[4] 创建扫描输出目录...")
    os.makedirs(SCAN_OUTPUT_DIR, exist_ok=True)
    print(f"    输出目录: {SCAN_OUTPUT_DIR}")
    
    # 5. 执行扫描命令
    print(f"\n[5] 执行扫描命令...")
    print(f"    输出文件: {SCAN_OUTPUT_PATH}")
    print(f"    提示：请确保扫描仪已连接并准备好扫描")
    
    # 构建PowerShell命令
    ps_command = f'''
$naps2ConsolePath = "{naps2_path}"
if (-not (Test-Path -Path $naps2ConsolePath -PathType Leaf)) {{
    Write-Error "错误：未找到NAPS2控制台程序"
    exit 1
}}
& $naps2ConsolePath -o "{SCAN_OUTPUT_PATH}"
'''
    
    try:
        print(f"\n    正在执行扫描（请等待，可能需要几十秒）...")
        result = subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_command],
            capture_output=True,
            text=True,
            timeout=120  # 120秒超时
        )
        
        if result.returncode != 0:
            print(f"    ✗ 扫描失败！")
            print(f"    返回码: {result.returncode}")
            if result.stderr:
                print(f"    错误信息: {result.stderr}")
            if result.stdout:
                print(f"    输出信息: {result.stdout}")
            return False
        
        print(f"    ✓ 扫描命令执行完成")
        if result.stdout:
            print(f"    输出: {result.stdout}")
        
        # 6. 检查输出文件
        print(f"\n[6] 检查输出文件...")
        if os.path.exists(SCAN_OUTPUT_PATH):
            file_size = os.path.getsize(SCAN_OUTPUT_PATH)
            print(f"    ✓ 扫描成功！文件已生成")
            print(f"    文件路径: {SCAN_OUTPUT_PATH}")
            print(f"    文件大小: {file_size:,} 字节 ({file_size / 1024 / 1024:.2f} MB)")
        else:
            print(f"    ⚠ 扫描命令执行完成，但未找到输出文件")
            print(f"    提示：请检查扫描仪是否正常工作，或查看NAPS2的错误信息")
            return False
        
        print(f"\n" + "=" * 60)
        print(f"测试完成！扫描文件已保存到: {SCAN_OUTPUT_PATH}")
        print(f"=" * 60)
        return True
        
    except subprocess.TimeoutExpired:
        print(f"    ✗ 扫描超时！请检查扫描仪连接或重试")
        return False
    except Exception as e:
        print(f"    ✗ 扫描过程中出现异常: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == '__main__':
    print(f"\n扫描测试程序")
    print(f"当前时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"当前目录: {os.path.dirname(__file__)}\n")
    
    success = test_scan()
    
    if success:
        print(f"\n✓ 测试成功！")
    else:
        print(f"\n✗ 测试失败，请检查配置和扫描仪连接")

