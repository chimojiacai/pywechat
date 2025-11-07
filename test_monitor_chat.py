"""
微信4.0版本 - 监控聊天窗口消息
功能：
1. 将某个聊天窗口拖出为独立窗口
2. 实时（每1秒）获取聊天窗口的消息
3. 返回消息数据
"""

from pyweixin.WeChatTools import Navigator, Tools
from pyweixin.Config import GlobalConfig
from pyweixin.Uielements import Main_window, Buttons, SideBar
from pyweixin.Errors import NoSuchFriendError, WeChatNotStartError
from pyweixin.WinSettings import Systemsettings
from pywinauto import Desktop
from pywinauto.controls.uia_controls import ListItemWrapper
import time
import re
import pyautogui
from typing import List, Dict, Optional, Tuple

# UI元素实例化
Main_window = Main_window()
Buttons = Buttons()
SideBar = SideBar()


def open_independent_chat_window(friend: str, is_maximize: bool = None) -> Desktop.window:
    """
    打开独立聊天窗口（将聊天窗口拖出）
    
    Args:
        friend: 好友或群聊备注名称
        is_maximize: 微信界面是否全屏，默认不全屏
        
    Returns:
        chat_window: 独立聊天窗口对象
    """
    if is_maximize is None:
        is_maximize = GlobalConfig.is_maximize
    
    # 1. 打开微信主窗口
    main_window = Navigator.open_weixin(is_maximize=is_maximize)
    
    # 2. 确保在聊天界面
    chats_button = main_window.child_window(**SideBar.Chats)
    message_list_pane = main_window.child_window(**Main_window.ConversationList)
    
    if not message_list_pane.exists():
        chats_button.click_input()
    if not message_list_pane.is_visible():
        chats_button.click_input()
    
    # 3. 先检查独立窗口是否已经存在（尝试多种方式）
    desktop = Desktop(backend='uia')
    
    # 方式1: 通过标题查找（不指定class_name，更通用）
    try:
        existing_window = desktop.window(title=friend)
        if existing_window.exists(timeout=0.5):
            print(f"✓ 发现已存在的独立窗口（通过标题）")
            Tools.cancel_pin(main_window)
            return existing_window
    except:
        pass
    
    # 方式2: 通过ChatWnd类名查找
    try:
        existing_window = desktop.window(title=friend, class_name='ChatWnd', framework_id='Win32')
        if existing_window.exists(timeout=0.5):
            print(f"✓ 发现已存在的独立窗口（ChatWnd）")
            Tools.cancel_pin(main_window)
            return existing_window
    except:
        pass
    
    # 4. 查找好友并打开独立窗口
    window_opened = False
    
    # 方式1: 检查是否已经在当前聊天窗口
    current_chat = main_window.child_window(control_type='Text', title=friend, found_index=0)
    if current_chat.exists():
        print(f"✓ 当前已在 {friend} 的聊天窗口")
        # 已经在当前聊天窗口，直接双击会话列表项
        message_list = message_list_pane.children(control_type='ListItem')
        for item in message_list:
            if friend in item.automation_id():
                print(f"  双击会话列表项打开独立窗口...")
                item.double_click_input()
                window_opened = True
                break
    else:
        # 方式2: 在会话列表中查找
        print(f"  在会话列表中查找 {friend}...")
        message_list = message_list_pane.children(control_type='ListItem')
        for item in message_list:
            if friend in item.automation_id():
                print(f"  找到会话列表项，双击打开独立窗口...")
                item.double_click_input()
                window_opened = True
                break
        
        # 方式3: 通过搜索栏查找
        if not window_opened:
            print(f"  通过搜索栏查找 {friend}...")
            try:
                search = main_window.descendants(**Main_window.Search)[0]
                if search.exists():
                    search.click_input()
                    
                    Systemsettings.copy_text_to_windowsclipboard(friend)
                    pyautogui.hotkey('ctrl', 'v')
                    
                    time.sleep(1)
                    search_results = main_window.child_window(title='', control_type='List')
                    if search_results.exists():
                        list_items = search_results.children(control_type="ListItem")
                        for item in list_items:
                            # 检查是否包含好友名称（去除空格等特殊字符）
                            item_text = item.window_text().replace('\u2002', ' ').replace('\u2004', ' ').replace('\u2005', ' ').replace('\u2006', ' ').replace('\u2009', ' ')
                            if friend in item_text:
                                print(f"  找到搜索结果，点击并打开独立窗口...")
                                item.click_input()
                                time.sleep(0.5)
                                
                                # 双击会话列表中的选中项
                                selected_items = [item for item in message_list_pane.children(control_type='ListItem') 
                                                if item.is_selected()]
                                if selected_items:
                                    selected_items[0].double_click_input()
                                    window_opened = True
                                else:
                                    # 如果没找到选中的，尝试通过automation_id查找
                                    for item in message_list_pane.children(control_type='ListItem'):
                                        if friend in item.automation_id():
                                            item.double_click_input()
                                            window_opened = True
                                            break
                                break
            except Exception as e:
                print(f"  搜索过程出错: {e}，继续尝试定位窗口...")
    
    # 5. 等待独立窗口出现并定位（给足够的时间让窗口打开）
    print(f"  等待独立窗口出现...")
    time.sleep(2)  # 增加等待时间
    
    # 尝试多种方式定位窗口
    chat_window = None
    max_retries = 5
    retry_count = 0
    
    while retry_count < max_retries and (not chat_window or not chat_window.exists()):
        retry_count += 1
        if retry_count > 1:
            print(f"  重试定位窗口 (第 {retry_count} 次)...")
            time.sleep(1)
        
        # 方式1: 只通过标题查找（最通用，不依赖class_name）
        try:
            chat_window = desktop.window(title=friend)
            if chat_window.exists(timeout=1):
                # 验证窗口确实存在且可见
                if chat_window.is_visible():
                    print(f"✓ 成功定位独立窗口（方式1: 标题匹配）")
                    break
        except Exception as e:
            pass
        
        # 方式1b: 通过标题和ChatWnd类名精确匹配
        try:
            chat_window = desktop.window(title=friend, class_name='ChatWnd', framework_id='Win32')
            if chat_window.exists(timeout=1):
                print(f"✓ 成功定位独立窗口（方式1b: ChatWnd精确匹配）")
                break
        except Exception as e:
            pass
        
        # 方式2: 查找所有窗口，通过标题匹配（不限制class_name）
        try:
            # 先尝试查找所有可见窗口，匹配标题
            all_windows = desktop.windows()
            for win in all_windows:
                try:
                    if win.exists(timeout=0.2) and win.is_visible():
                        win_text = win.window_text()
                        # 检查窗口标题是否完全匹配或包含好友名称
                        if win_text == friend or friend in win_text:
                            # 验证窗口确实可见
                            chat_window = win
                            if chat_window.exists(timeout=0.5) and chat_window.is_visible():
                                print(f"✓ 成功定位独立窗口（方式2: 全窗口搜索，窗口标题: {win_text}）")
                                break
                except:
                    continue
            if chat_window and chat_window.exists() and chat_window.is_visible():
                break
        except Exception as e:
            pass
        
        # 方式2b: 只通过ChatWnd类名查找，然后匹配标题内容
        try:
            all_windows = desktop.windows(class_name='ChatWnd', framework_id='Win32')
            for win in all_windows:
                try:
                    if win.exists(timeout=0.3):
                        win_text = win.window_text()
                        win_name = win.element_info.name if hasattr(win.element_info, 'name') else ''
                        # 检查窗口标题或名称中是否包含好友名称
                        if friend in win_text or friend in win_name:
                            chat_window = win
                            if chat_window.exists(timeout=0.5):
                                print(f"✓ 成功定位独立窗口（方式2b: ChatWnd标题匹配，窗口标题: {win_text}）")
                                break
                except:
                    continue
            if chat_window and chat_window.exists():
                break
        except Exception as e:
            pass
        
        # 方式3: 查找所有ChatWnd窗口，取最新的（通常是最新打开的）
        try:
            all_windows = desktop.windows(class_name='ChatWnd', framework_id='Win32')
            if all_windows:
                # 取最后一个（通常是最新打开的）
                chat_window = all_windows[-1]
                if chat_window.exists(timeout=0.5) and chat_window.is_visible():
                    try:
                        win_text = chat_window.window_text()
                        print(f"✓ 成功定位独立窗口（方式3: 最新ChatWnd窗口，窗口标题: {win_text}）")
                    except:
                        print(f"✓ 成功定位独立窗口（方式3: 最新ChatWnd窗口）")
                    break
        except Exception as e:
            pass
        
        # 方式3b: 查找所有可见窗口，取标题匹配的最新窗口
        try:
            all_windows = desktop.windows()
            matching_windows = []
            for win in all_windows:
                try:
                    if win.exists(timeout=0.2) and win.is_visible():
                        win_text = win.window_text()
                        if friend in win_text:
                            matching_windows.append(win)
                except:
                    continue
            if matching_windows:
                # 取最后一个匹配的窗口
                chat_window = matching_windows[-1]
                if chat_window.exists(timeout=0.5):
                    try:
                        win_text = chat_window.window_text()
                        print(f"✓ 成功定位独立窗口（方式3b: 最新匹配窗口，窗口标题: {win_text}）")
                    except:
                        print(f"✓ 成功定位独立窗口（方式3b: 最新匹配窗口）")
                    break
        except Exception as e:
            pass
        
        # 方式4: 尝试通过部分标题匹配
        try:
            all_windows = desktop.windows(class_name='ChatWnd', framework_id='Win32')
            for win in all_windows:
                try:
                    if win.exists(timeout=0.3):
                        win_text = win.window_text()
                        # 如果窗口标题包含好友名称的部分字符，也认为匹配
                        if len(friend) >= 3 and any(part in win_text for part in [friend[:3], friend[-3:]]):
                            chat_window = win
                            print(f"✓ 成功定位独立窗口（方式4: 部分匹配，窗口标题: {win_text}）")
                            break
                except:
                    continue
            if chat_window and chat_window.exists():
                break
        except Exception as e:
            pass
    
    # 如果还是找不到，尝试最后一次：打印所有窗口信息用于调试
    if not chat_window or not chat_window.exists():
        print(f"  调试信息: 尝试查找所有窗口...")
        try:
            # 打印所有ChatWnd窗口
            all_chat_wnd = desktop.windows(class_name='ChatWnd', framework_id='Win32')
            print(f"    找到 {len(all_chat_wnd)} 个ChatWnd窗口:")
            for i, win in enumerate(all_chat_wnd):
                try:
                    if win.exists(timeout=0.2):
                        win_text = win.window_text()
                        class_name = win.element_info.class_name if hasattr(win.element_info, 'class_name') else '未知'
                        print(f"      窗口{i+1}: 标题='{win_text}', class_name='{class_name}'")
                except Exception as e:
                    print(f"      窗口{i+1}: 无法获取信息 - {e}")
        except Exception as e:
            print(f"    查找ChatWnd窗口出错: {e}")
        
        try:
            # 打印所有包含好友名称的窗口
            all_windows = desktop.windows()
            matching = []
            for win in all_windows:
                try:
                    if win.exists(timeout=0.1) and win.is_visible():
                        win_text = win.window_text()
                        if friend in win_text:
                            class_name = win.element_info.class_name if hasattr(win.element_info, 'class_name') else '未知'
                            matching.append((win, win_text, class_name))
                except:
                    continue
            if matching:
                print(f"    找到 {len(matching)} 个标题包含'{friend}'的窗口:")
                for i, (win, win_text, class_name) in enumerate(matching):
                    print(f"      窗口{i+1}: 标题='{win_text}', class_name='{class_name}'")
        except Exception as e:
            print(f"    查找所有窗口出错: {e}")
        
        # 如果还是找不到，但窗口可能已经打开了，尝试使用最新的匹配窗口
        try:
            all_windows = desktop.windows()
            matching_windows = []
            for win in all_windows:
                try:
                    if win.exists(timeout=0.2) and win.is_visible():
                        win_text = win.window_text()
                        if friend in win_text:
                            matching_windows.append(win)
                except:
                    continue
            if matching_windows:
                chat_window = matching_windows[-1]
                if chat_window.exists(timeout=0.5):
                    try:
                        win_text = chat_window.window_text()
                        print(f"✓ 使用最新匹配的窗口（窗口标题: {win_text}）")
                    except:
                        print(f"✓ 使用最新匹配的窗口")
                else:
                    chat_window = None
        except:
            pass
    
    # 如果还是找不到，抛出异常
    if not chat_window or not chat_window.exists():
        raise NoSuchFriendError(f"无法打开或定位 {friend} 的独立聊天窗口。请确认窗口是否已打开。")
    
    # 取消主窗口置顶
    Tools.cancel_pin(main_window)
    
    return chat_window


def parse_message_item(list_item: ListItemWrapper, is_group: bool = False) -> Dict[str, str]:
    """
    解析单条消息ListItem
    
    Args:
        list_item: 消息ListItem对象
        is_group: 是否为群聊
        
    Returns:
        dict: 包含sender, content, type的消息字典
    """
    message = {
        'sender': '',
        'content': '',
        'type': '未知',
        'raw_text': list_item.window_text()
    }
    
    # 检查是否有按钮（有按钮的是用户消息，没有的是系统消息）
    buttons = list_item.descendants(control_type='Button')
    
    if len(buttons) == 0:
        # 系统消息
        message['sender'] = '系统'
        message['content'] = list_item.window_text()
        message['type'] = '系统消息'
    else:
        # 用户消息
        try:
            # 获取发送者（通常是第一个按钮的文本）
            if buttons:
                message['sender'] = buttons[0].window_text()
        except:
            message['sender'] = '未知'
        
        # 获取消息内容
        raw_text = list_item.window_text()
        
        # 特殊消息类型判断（简体中文）
        special_messages = {
            '[图片]': '图片',
            '[视频]': '视频',
            '[动画表情]': '动画表情',
            '[视频号]': '视频号',
            '[链接]': '链接',
            '[聊天记录]': '聊天记录',
            '[文件]': '文件'
        }
        
        if raw_text in special_messages:
            message['type'] = special_messages[raw_text]
            message['content'] = special_messages[raw_text]
        elif raw_text == '[文件]':
            # 文件消息，尝试获取文件名
            texts = list_item.descendants(control_type='Text')
            if texts:
                filename = texts[0].window_text()
                message['content'] = f'文件: {filename}'
                message['type'] = '文件'
        elif re.match(r'\[语音\]\d+秒', raw_text):
            # 语音消息
            message['type'] = '语音'
            # 尝试获取语音转文字
            try:
                texts = list_item.descendants(control_type='Text')
                if is_group and len(texts) >= 3:
                    audio_content = texts[2].window_text()
                    message['content'] = f'{raw_text} 转文字: {audio_content}'
                elif len(texts) >= 2:
                    audio_content = texts[1].window_text()
                    message['content'] = f'{raw_text} 转文字: {audio_content}'
                else:
                    message['content'] = raw_text
            except:
                message['content'] = raw_text
        else:
            # 普通文本消息
            message['type'] = '文本'
            message['content'] = raw_text
    
    return message


def get_chat_messages(chat_window, max_messages: int = 50, chats_only: bool = True) -> List[Dict[str, str]]:
    """
    从独立聊天窗口获取消息
    
    Args:
        chat_window: 独立聊天窗口对象
        max_messages: 最多获取多少条消息
        chats_only: 是否只获取用户消息（排除系统消息）
        
    Returns:
        list: 消息列表，每个元素是包含sender, content, type的字典
    """
    messages = []
    
    try:
        # 确保窗口可见且激活
        if not chat_window.exists():
            return messages
        
        # 激活窗口（确保窗口在前台）
        try:
            chat_window.set_focus()
            time.sleep(0.1)
        except:
            pass
        
        # 定位消息列表（尝试多种方式）
        chat_list = None
        
        # 方式1: 通过FriendChatList定位
        try:
            chat_list = chat_window.child_window(**Main_window.FriendChatList)
            if not chat_list.exists(timeout=0.5):
                chat_list = None
        except:
            pass
        
        # 方式2: 通过title='消息'定位
        if not chat_list or not chat_list.exists():
            try:
                chat_list = chat_window.child_window(title='消息', control_type='List')
                if not chat_list.exists(timeout=0.5):
                    chat_list = None
            except:
                pass
        
        # 方式3: 查找所有List控件，找到消息列表
        if not chat_list or not chat_list.exists():
            try:
                all_lists = chat_window.descendants(control_type='List')
                for lst in all_lists:
                    try:
                        if lst.exists(timeout=0.2):
                            # 检查是否包含ListItem（消息列表的特征）
                            items = lst.children(control_type='ListItem')
                            if len(items) > 0:
                                # 验证是否是消息列表（通常消息列表会有很多项）
                                chat_list = lst
                                break
                    except:
                        continue
            except:
                pass
        
        if not chat_list or not chat_list.exists():
            # 调试信息
            try:
                all_lists = chat_window.descendants(control_type='List')
                print(f"  调试: 找到 {len(all_lists)} 个List控件")
                for i, lst in enumerate(all_lists):
                    try:
                        if lst.exists(timeout=0.1):
                            items = lst.children(control_type='ListItem')
                            print(f"    List{i+1}: {len(items)} 个ListItem")
                    except:
                        pass
            except:
                pass
            return messages
        
        # 点击消息区域激活（确保消息列表可访问）
        try:
            # 点击消息列表区域左侧，激活滚动条
            rect = chat_list.rectangle()
            x = rect.left + 8
            y = (rect.top + rect.bottom) // 2
            from pywinauto import mouse
            mouse.click(coords=(x, y))
            time.sleep(0.1)
        except:
            pass
        
        # 滚动到底部（确保看到最新消息）
        try:
            chat_list.type_keys('{END}')
            time.sleep(0.2)
            # 再按一次确保到底
            chat_list.type_keys('{END}')
            time.sleep(0.1)
        except:
            pass
        
        # 判断是否为群聊（检查是否有视频通话按钮）
        try:
            video_call_button = chat_window.child_window(**Buttons.VideoCallButton)
            is_group = not video_call_button.exists(timeout=0.3)
        except:
            is_group = False
        
        # 获取所有消息ListItem
        try:
            list_items = chat_list.children(control_type='ListItem')
        except:
            return messages
        
        if len(list_items) == 0:
            return messages
        
        # 过滤掉"查看更多消息"按钮
        check_more_title = Buttons.CheckMoreMessagesButton.get('title', '查看更多消息')
        list_items = [item for item in list_items 
                      if item.window_text() != check_more_title]
        
        # 如果只要用户消息，过滤掉系统消息
        if chats_only:
            list_items = [item for item in list_items 
                         if len(item.descendants(control_type='Button')) > 0]
        
        # 限制数量（取最新的消息）
        if len(list_items) > max_messages:
            list_items = list_items[-max_messages:]
        
        # 解析每条消息
        for item in list_items:
            try:
                message = parse_message_item(item, is_group=is_group)
                messages.append(message)
            except Exception as e:
                # 解析失败，记录原始文本
                try:
                    raw_text = item.window_text()
                except:
                    raw_text = '无法获取文本'
                messages.append({
                    'sender': '解析失败',
                    'content': raw_text,
                    'type': '未知',
                    'raw_text': raw_text,
                    'error': str(e)
                })
    
    except Exception as e:
        print(f"  获取消息时出错: {e}")
        import traceback
        traceback.print_exc()
    
    return messages


def monitor_chat_window(friend: str, 
                       interval: float = 1.0, 
                       duration: Optional[float] = None,
                       callback: Optional[callable] = None) -> List[Dict]:
    """
    监控聊天窗口消息
    
    Args:
        friend: 好友或群聊备注名称
        interval: 监控间隔（秒），默认1秒
        duration: 监控持续时间（秒），None表示持续监控直到手动停止
        callback: 回调函数，每次获取到新消息时调用 callback(new_messages)
        
    Returns:
        list: 所有获取到的消息列表
    """
    print(f"正在打开 {friend} 的独立聊天窗口...")
    
    # 1. 打开独立窗口
    chat_window = open_independent_chat_window(friend)
    print(f"✓ 独立聊天窗口已打开")
    
    # 2. 获取初始消息（作为基准）
    print(f"  正在获取初始消息...")
    previous_messages = get_chat_messages(chat_window, max_messages=100, chats_only=False)
    previous_count = len(previous_messages)
    print(f"✓ 初始消息数量: {previous_count} 条")
    if previous_count > 0:
        print(f"  最新消息示例: {previous_messages[-1].get('content', '')[:50]}")
    
    all_messages = previous_messages.copy()
    start_time = time.time()
    
    print(f"\n开始监控... (间隔: {interval}秒)")
    if duration:
        print(f"监控时长: {duration}秒")
    print("按 Ctrl+C 停止监控\n")
    
    try:
        while True:
            # 检查是否超时
            if duration and (time.time() - start_time) >= duration:
                print(f"\n监控时间已到 ({duration}秒)")
                break
            
            # 获取当前消息
            current_messages = get_chat_messages(chat_window, max_messages=100, chats_only=False)
            current_count = len(current_messages)
            
            # 检测新消息（通过消息内容对比，更可靠）
            if current_count > previous_count:
                # 方式1: 通过数量变化检测
                new_count = current_count - previous_count
                new_messages = current_messages[-new_count:]
            else:
                # 方式2: 通过内容对比检测（处理消息被删除或更新的情况）
                new_messages = []
                if previous_count > 0 and current_count > 0:
                    # 获取之前的最后几条消息的原始文本作为指纹
                    prev_fingerprints = [msg.get('raw_text', '') for msg in previous_messages[-10:]]
                    # 检查当前消息中是否有新的
                    for msg in current_messages:
                        msg_fingerprint = msg.get('raw_text', '')
                        if msg_fingerprint and msg_fingerprint not in prev_fingerprints:
                            new_messages.append(msg)
            
            if new_messages:
                print(f"[{time.strftime('%H:%M:%S')}] 检测到 {len(new_messages)} 条新消息:")
                for msg in new_messages:
                    sender = msg.get('sender', '未知')
                    content = msg.get('content', '')[:50]  # 限制显示长度
                    msg_type = msg.get('type', '未知')
                    print(f"  [{msg_type}] {sender}: {content}")
                
                # 调用回调函数
                if callback:
                    try:
                        callback(new_messages)
                    except Exception as e:
                        print(f"  回调函数执行错误: {e}")
                
                # 更新消息列表
                all_messages.extend(new_messages)
            
            previous_messages = current_messages
            previous_count = current_count
            
            # 等待下次检查
            time.sleep(interval)
            
    except KeyboardInterrupt:
        print("\n\n监控已停止")
    except Exception as e:
        print(f"\n监控出错: {e}")
        import traceback
        traceback.print_exc()
    
    print(f"\n总共获取到 {len(all_messages)} 条消息")
    return all_messages


if __name__ == "__main__":
    try:
        # 配置
        GlobalConfig.search_pages = 1
        
        # 要监控的好友或群聊
        friend = "0xlyz2"  # 修改为你要监控的好友名称
        
        # 定义回调函数（可选）
        def on_new_messages(new_messages):
            """新消息回调函数"""
            for msg in new_messages:
                # 这里可以处理新消息，比如保存到数据库、发送通知等
                pass
        
        # 开始监控
        # 参数说明：
        # - friend: 好友或群聊名称
        # - interval: 检查间隔（秒），默认1秒
        # - duration: 监控时长（秒），None表示持续监控
        # - callback: 新消息回调函数（可选）
        
        messages = monitor_chat_window(
            friend=friend,
            interval=1.0,  # 每1秒检查一次
            duration=None,  # None表示持续监控，可以设置为具体秒数，如 60 表示监控60秒
            callback=on_new_messages  # 可选的回调函数
        )
        
        # 打印所有消息
        print("\n" + "="*60)
        print("所有消息:")
        print("="*60)
        for i, msg in enumerate(messages, 1):
            print(f"{i}. [{msg['type']}] {msg['sender']}: {msg['content']}")
        
    except NoSuchFriendError as e:
        print(f"\n✗ 错误: 未找到好友")
        print(f"   好友名称: {friend}")
        print("\n解决方法:")
        print("1. 检查好友备注名称是否正确（必须与微信中显示的一致）")
        print("2. 建议先用'文件传输助手'测试")
    except WeChatNotStartError as e:
        print(f"\n✗ 错误: 微信启动失败")
        print(f"   {str(e)}")
    except Exception as e:
        print(f"\n✗ 错误: {type(e).__name__}")
        print(f"   错误信息: {str(e)}")
        import traceback
        traceback.print_exc()

