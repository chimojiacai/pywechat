"""
微信4.0版本 - 快速发送文件测试（简化版）
最简单的测试示例 - 只负责调用方法，窗口等待由pyweixin库处理
"""

from pyweixin.WeChatAuto import Files
from pyweixin.Config import GlobalConfig
from pyweixin.Errors import (
    NoSuchFriendError,
    NoFilesToSendError,
    NetWorkNotConnectError,
    ScanCodeToLogInError,
    NotInstalledError,
    WeChatNotStartError
)
import os

if __name__ == "__main__":
    try:
        print("="*60)
        print("微信4.0版本 - 发送文件测试")
        print("="*60)
        
        # 配置：使用搜索栏查找好友（更可靠）
        GlobalConfig.search_pages = 1
        
        # 发送文件 - 使用 send_files_to_friend 方法（单发）
        friend = "文件传输助手"  # 修改为你要测试的好友名称
        
        # 文件路径列表 - 可以发送多个文件
        # 注意：文件路径必须是绝对路径，且文件必须存在
        files = [
            # 示例：修改为你要发送的文件路径
            r"C:\Users\liyongzhen\work\chatlog\main.go",
            # r"C:\Users\YourName\Pictures\image.jpg",
        ]
        
        # 可选：发送文件时同时发送消息
        with_messages = False  # 设置为True可以同时发送消息
        messages = [
            "这是随文件一起发送的消息 📎",
            # "第二条消息",
        ]
        messages_first = False  # True: 先发消息后发文件, False: 先发文件后发消息
        
        # 检查文件是否存在
        if not files:
            print("\n⚠ 警告: 文件列表为空！")
            print("请在代码中修改files列表，添加要发送的文件路径（绝对路径）")
            print("\n示例:")
            print('  files = [')
            print('      r"C:\\Users\\YourName\\Documents\\test.txt",')
            print('      r"C:\\Users\\YourName\\Pictures\\image.jpg",')
            print('  ]')
            exit(1)
        
        # 验证文件是否存在
        valid_files = []
        for file_path in files:
            if os.path.isfile(file_path):
                file_size = os.path.getsize(file_path)
                if 0 < file_size < 1073741824:  # 文件大小必须在0到1GB之间
                    valid_files.append(file_path)
                    print(f"✓ 文件有效: {os.path.basename(file_path)} ({file_size / 1024 / 1024:.2f} MB)")
                else:
                    print(f"✗ 文件大小无效: {os.path.basename(file_path)} (大小: {file_size / 1024 / 1024:.2f} MB)")
            else:
                print(f"✗ 文件不存在: {file_path}")
        
        if not valid_files:
            print("\n✗ 错误: 没有有效的文件可以发送")
            print("请检查:")
            print("1. 文件路径是否正确（必须是绝对路径）")
            print("2. 文件是否存在")
            print("3. 文件大小是否在0到1GB之间")
            exit(1)
        
        print(f"\n准备发送文件给: {friend}")
        print(f"文件数量: {len(valid_files)} 个")
        for i, file_path in enumerate(valid_files, 1):
            print(f"  文件{i}: {os.path.basename(file_path)}")
        if with_messages:
            print(f"同时发送消息: {len(messages)} 条")
            for i, msg in enumerate(messages, 1):
                print(f"  消息{i}: {msg}")
        print()
        
        Files.send_files_to_friend(
            friend=friend,
            files=valid_files,
            with_messages=with_messages,
            messages=messages if with_messages else [],
            messages_first=messages_first,
            send_delay=1,
            is_maximize=False,
            close_weixin=True
        )
        
        print("\n" + "="*60)
        print("✓ 文件发送成功！")
        print("="*60)
        
    except NoSuchFriendError as e:
        print(f"\n✗ 错误: 未找到好友")
        try:
            print(f"   好友名称: {friend}")
        except:
            print("   好友名称: 未定义")
        print("\n解决方法:")
        print("1. 检查好友备注名称是否正确（必须与微信中显示的一致）")
        print("2. 建议先用'文件传输助手'测试")
        print("3. 确保好友在通讯录中")
    except NoFilesToSendError as e:
        print(f"\n✗ 错误: 没有可发送的文件")
        print("   请检查:")
        print("   1. 文件路径是否正确（必须是绝对路径）")
        print("   2. 文件是否存在且可读")
        print("   3. 文件大小是否在0到1GB之间")
        print("   4. 文件是否为空")
    except NetWorkNotConnectError as e:
        print(f"\n✗ 错误: 网络连接问题")
        print("   请检查:")
        print("   1. 网络连接是否正常")
        print("   2. 微信是否可以正常联网")
    except ScanCodeToLogInError as e:
        print(f"\n✗ 错误: 需要扫码登录")
        print("   解决方法:")
        print("   1. 确保微信已登录")
        print("   2. 在手机端开启'PC端自动登录'功能")
    except NotInstalledError as e:
        print(f"\n✗ 错误: 未找到微信")
        print("   请确保已安装微信4.0版本")
    except WeChatNotStartError as e:
        print(f"\n✗ 错误: 微信启动失败")
        print(f"   {str(e)}")
        print("\n解决方法:")
        print("1. 手动打开微信并确保已登录")
        print("2. 确保微信窗口可见（不要最小化）")
        print("3. 如果使用微信4.0，可能需要开启讲述人模式")
    except Exception as e:
        print(f"\n✗ 发送失败: {type(e).__name__}")
        print(f"   错误信息: {str(e)}")
        print("\n常见问题:")
        print("1. 确保微信已登录")
        print("2. 确保好友名称正确（建议使用'文件传输助手'测试）")
        print("3. 确保网络连接正常")
        print("4. 检查微信版本是否为4.0+")
        print("5. 确保文件路径是绝对路径且文件存在")
        print("6. 确保文件大小在0到1GB之间")
        print("7. 如果使用微信4.0，可能需要开启讲述人模式（查看Weixin4.0.md）")
        
        # 打印详细错误信息用于调试
        import traceback
        print("\n详细错误信息:")
        traceback.print_exc()

