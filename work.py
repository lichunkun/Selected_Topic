import win32api
import win32gui
import win32con
import win32clipboard as w


def getText():
    """获取剪贴板文本"""
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_UNICODETEXT)
    w.CloseClipboard()
    print(d)
    return d

def setText(aString):
    """设置剪贴板文本"""
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
    w.CloseClipboard()

def backPycharm():
    # 返回pycharm
    # target = u'*Python 3.5.4 Shell*'
    # pycharm = win32gui.FindWindow(None, target)
    # win32gui.SendMessage(pycharm,0,0,0)
    win32api.SetCursorPos([1500, 800])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

flag = 1
number_list = []
while flag:
    # getText()
    try:
        number = int(input("请输入论文编号:"))
    except:
        number = None
        print("请输入正确的数字")
        continue
    if number in number_list:
        print("_____%s号论文已被选择_____"%(number))
        # QQ群发送消息
        to_who1 = u'赵晓月'  # 接收消息qq的备注名称(该好友对话框单独打开，最小化)
        content = u"_____%s号论文已被选择_____"%(number)  # 要发送的消息
        setText(content)
        qqhd = win32gui.FindWindow(None, to_who1)
        # 投递剪贴板消息到QQ窗体
        win32gui.SendMessage(qqhd, 258, 22, 2080193)
        win32gui.SendMessage(qqhd, 770, 0, 0)
        # 模拟按下回车键
        win32gui.SendMessage(qqhd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.SendMessage(qqhd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        backPycharm()
        continue
    number_list.append(number)

    # print(number_list)
