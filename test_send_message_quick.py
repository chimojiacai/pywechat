"""
微信4.0版本 - 快速发送消息测试（简化版）
最简单的测试示例 - 只负责调用方法，窗口等待由pyweixin库处理
"""

from pyweixin.WeChatAuto import Messages
from pyweixin.Config import GlobalConfig
from pyweixin.Errors import (
    NoSuchFriendError,
    CantSendEmptyMessageError,
    NetWorkNotConnectError,
    ScanCodeToLogInError,
    NotInstalledError,
    WeChatNotStartError
)

if __name__ == "__main__":
    try:
        print("="*60)
        print("微信4.0版本 - 发送消息测试")
        print("="*60)
        
        # 配置：使用搜索栏查找好友（更可靠）
        GlobalConfig.search_pages = 1
        
        # 发送消息 - 使用 send_messages_to_friend 方法（单发）
        friend = "探活ping"  # 修改为你要测试的好友名称
        messages = [
            "Hello! 这是来自pyweixin的测试消息 🎉",  # 可以发送多条消息
            # "第二条消息",
            # "第三条消息"
        ]
        
        print(f"\n准备发送消息给: {friend}")
        print(f"消息数量: {len(messages)} 条")
        for i, msg in enumerate(messages, 1):
            print(f"  消息{i}: {msg}")
        print()
        
        Messages.send_messages_to_friend(
            friend=friend,
            messages=messages,
            send_delay=1,
            is_maximize=False,
            close_weixin=True
        )
        
        print("\n" + "="*60)
        print("✓ 消息发送成功！")
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
    except CantSendEmptyMessageError as e:
        print(f"\n✗ 错误: 消息内容为空")
        print("   请确保messages列表不为空，且每条消息都有内容")
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
        print("5. 如果使用微信4.0，可能需要开启讲述人模式（查看Weixin4.0.md）")
        
        # 打印详细错误信息用于调试
        import traceback
        print("\n详细错误信息:")
        traceback.print_exc()

