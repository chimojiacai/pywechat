"""
微信4.0版本诊断脚本
用于检查微信环境和窗口查找问题
"""

from pyweixin.WeChatTools import Navigator, Tools
from pyweixin.Errors import NotInstalledError
from pywinauto import Desktop
from pyweixin.Uielements import Main_window, Login_window

def diagnose():
    print("="*60)
    print("微信4.0版本 - 环境诊断")
    print("="*60)
    
    # 1. 检查微信是否运行
    print("\n[1] 检查微信运行状态...")
    is_running = Tools.is_weixin_running()
    print(f"   结果: {'✓ 微信正在运行' if is_running else '✗ 微信未运行'}")
    
    # 2. 检查微信路径
    print("\n[2] 检查微信安装路径...")
    try:
        weixin_path = Tools.find_weixin_path(copy_to_clipboard=False)
        print(f"   ✓ 微信路径: {weixin_path}")
    except NotInstalledError:
        print("   ✗ 未找到微信安装路径")
        print("   请确保已安装微信4.0版本")
        return False
    except Exception as e:
        print(f"   ✗ 查找路径失败: {str(e)}")
        return False
    
    # 3. 尝试查找登录窗口
    print("\n[3] 检查登录窗口...")
    try:
        desktop = Desktop(backend='uia')
        login_window = desktop.window(**Login_window().LoginWindow)
        if login_window.exists(timeout=2):
            print("   ✓ 找到登录窗口（微信未登录）")
        else:
            print("   ✗ 未找到登录窗口")
    except Exception as e:
        print(f"   ⚠ 检查登录窗口时出错: {str(e)}")
    
    # 4. 尝试查找主窗口
    print("\n[4] 检查主窗口...")
    try:
        desktop = Desktop(backend='uia')
        main_window_spec = Main_window().MainWindow
        print(f"   窗口规格: {main_window_spec}")
        
        main_window = desktop.window(**main_window_spec)
        print(f"   窗口对象类型: {type(main_window)}")
        
        if main_window.exists(timeout=2):
            print("   ✓ 找到主窗口（微信已登录）")
            print(f"   窗口句柄: {main_window.handle}")
            print(f"   窗口标题: {main_window.window_text()}")
            rect = main_window.rectangle()
            print(f"   窗口位置: ({rect.left}, {rect.top}), 大小: {rect.width()}x{rect.height()}")
        else:
            print("   ✗ 未找到主窗口")
            print("   可能的原因:")
            print("   1. 微信未登录")
            print("   2. 微信窗口被最小化或隐藏")
            print("   3. 微信版本不兼容")
            print("   4. 需要开启讲述人模式（查看Weixin4.0.md）")
    except Exception as e:
        print(f"   ✗ 检查主窗口时出错: {str(e)}")
        import traceback
        traceback.print_exc()
    
    # 5. 尝试打开微信
    print("\n[5] 测试打开微信...")
    try:
        main_window = Navigator.open_weixin(is_maximize=False)
        print(f"   返回对象类型: {type(main_window)}")
        
        if isinstance(main_window, dict):
            print("   ⚠ 警告: 返回的是字典而不是窗口对象")
            print("   这通常意味着窗口查找失败")
            print(f"   返回内容: {main_window}")
        elif hasattr(main_window, 'exists'):
            if main_window.exists(timeout=2):
                print("   ✓ 成功打开微信窗口")
                main_window.close()
            else:
                print("   ✗ 窗口打开失败（窗口不存在）")
        else:
            print(f"   ⚠ 未知的返回类型: {type(main_window)}")
    except Exception as e:
        print(f"   ✗ 打开微信失败: {str(e)}")
        import traceback
        traceback.print_exc()
    
    # 6. 检查wxid（需要微信已登录）
    print("\n[6] 检查当前登录账号...")
    try:
        wxid = Tools.get_current_wxid()
        if wxid:
            print(f"   ✓ 当前登录账号: {wxid}")
        else:
            print("   ✗ 无法获取wxid（可能未登录）")
    except Exception as e:
        print(f"   ⚠ 获取wxid失败: {str(e)}")
    
    print("\n" + "="*60)
    print("诊断完成")
    print("="*60)
    print("\n建议:")
    print("1. 确保微信已登录")
    print("2. 如果使用微信4.0，可能需要开启讲述人模式")
    print("3. 查看 Weixin4.0.md 了解详细说明")
    print("4. 确保微信版本为4.0.6+（推荐4.1.4.10）")

if __name__ == "__main__":
    diagnose()

