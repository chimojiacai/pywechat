"""
检查微信4.1.2.17版本的UI元素是否可访问
用于诊断版本兼容性问题
"""

from pywinauto import Desktop
from pyweixin.Uielements import Main_window, Login_window, SideBar
import time

def check_ui_elements():
    """检查关键UI元素是否可访问"""
    print("="*60)
    print("微信UI元素可访问性检查")
    print("="*60)
    
    desktop = Desktop(backend='uia')
    
    # 检查主窗口
    print("\n[1] 检查主窗口...")
    main_window_spec = Main_window().MainWindow
    print(f"   窗口规格: {main_window_spec}")
    
    main_window = desktop.window(**main_window_spec)
    if main_window.exists(timeout=3):
        print("   ✓ 主窗口可访问")
        print(f"   窗口句柄: {main_window.handle}")
        print(f"   窗口标题: {main_window.window_text()}")
        
        # 检查关键子元素
        print("\n[2] 检查主窗口内的关键元素...")
        
        # 检查搜索栏
        search_spec = Main_window().Search
        print(f"\n   搜索栏规格: {search_spec}")
        try:
            search = main_window.child_window(**search_spec)
            if search.exists(timeout=1):
                print("   ✓ 搜索栏可访问")
            else:
                print("   ✗ 搜索栏不可访问")
        except Exception as e:
            print(f"   ✗ 搜索栏检查失败: {str(e)}")
        
        # 检查会话列表
        conv_list_spec = Main_window().ConversationList
        print(f"\n   会话列表规格: {conv_list_spec}")
        try:
            conv_list = main_window.child_window(**conv_list_spec)
            if conv_list.exists(timeout=1):
                print("   ✓ 会话列表可访问")
                items = conv_list.children(control_type='ListItem')
                print(f"   会话列表项数量: {len(items)}")
            else:
                print("   ✗ 会话列表不可访问")
        except Exception as e:
            print(f"   ✗ 会话列表检查失败: {str(e)}")
        
        # 检查侧边栏
        print(f"\n   侧边栏规格: {SideBar().Chats}")
        try:
            chats_button = main_window.child_window(**SideBar().Chats)
            if chats_button.exists(timeout=1):
                print("   ✓ 聊天按钮可访问")
            else:
                print("   ✗ 聊天按钮不可访问")
        except Exception as e:
            print(f"   ✗ 聊天按钮检查失败: {str(e)}")
        
        # 尝试获取所有可访问的元素
        print("\n[3] 尝试获取主窗口的所有子元素...")
        try:
            all_children = main_window.descendants()
            print(f"   找到 {len(all_children)} 个子元素")
            
            # 按类型分类
            element_types = {}
            for child in all_children[:50]:  # 只显示前50个
                try:
                    ctrl_type = child.element_info.control_type
                    element_types[ctrl_type] = element_types.get(ctrl_type, 0) + 1
                except:
                    pass
            
            print("\n   元素类型统计（前50个）:")
            for ctrl_type, count in sorted(element_types.items()):
                print(f"     {ctrl_type}: {count}")
                
        except Exception as e:
            print(f"   ✗ 获取子元素失败: {str(e)}")
        
    else:
        print("   ✗ 主窗口不可访问")
        print("\n   可能的原因:")
        print("   1. 微信未登录")
        print("   2. 微信窗口被隐藏或最小化")
        print("   3. 需要开启讲述人模式（微信4.0+）")
        print("   4. UI元素定义与当前版本不匹配")
    
    # 检查登录窗口
    print("\n[4] 检查登录窗口...")
    login_window_spec = Login_window().LoginWindow
    print(f"   登录窗口规格: {login_window_spec}")
    login_window = desktop.window(**login_window_spec)
    if login_window.exists(timeout=1):
        print("   ✓ 登录窗口可访问（微信未登录）")
    else:
        print("   ✗ 登录窗口不可访问（可能已登录）")
    
    print("\n" + "="*60)
    print("检查完成")
    print("="*60)
    print("\n建议:")
    print("1. 如果主窗口不可访问，尝试开启讲述人模式")
    print("2. 如果元素不可访问，可能是版本不匹配，需要更新UI元素定义")
    print("3. 确保微信窗口可见（不要最小化）")

if __name__ == "__main__":
    try:
        check_ui_elements()
    except Exception as e:
        print(f"\n发生错误: {str(e)}")
        import traceback
        traceback.print_exc()

