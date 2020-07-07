
import wx
import pyautogui
import time
import os
from PIL import Image
import pyocr
import pyocr.builders
import win32gui
import datetime


# OCRエンジンの取得
tools = pyocr.get_available_tools()
tool = tools[0]

parent_handle = None

class MyFrame(wx.Frame):
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title, size=(450, 300))

        # キーの設定
        #[Alt]+[Space]
        new_id = wx.NewIdRef(count=1)
        self.RegisterHotKey(new_id, wx.MOD_ALT, wx.WXK_SPACE)  # Windows向け
        self.Bind(wx.EVT_HOTKEY, self.callhotkey,id=new_id)

        #[Ctrl]+[Space]
        new_id = wx.NewIdRef(count=1)
        self.RegisterHotKey(new_id, wx.MOD_CONTROL, wx.WXK_SPACE)  # Windows向け
        self.Bind(wx.EVT_HOTKEY, self.lockOn,id=new_id)
        
    def lockOn(self, event):
        """ ホットキーイベント """
        global parent_handle
        parent_handle = win32gui.GetForegroundWindow()
        
    def callhotkey(self, event):
        """ ホットキーイベント """
        global parent_handle
        if(parent_handle == None):
            print("親ハンドルなし")
            return

        yMargin = spinctrl.GetValue()		
        xMargin = spinctrl2.GetValue()        
        x2Margin = spinctrl3.GetValue()
        y2Margin = spinctrl4.GetValue()
		# スクリーンショット
        win_x1,win_y1,win_x2,win_y2 = win32gui.GetWindowRect(parent_handle)
        screenshot = pyautogui.screenshot(region=(win_x1 + xMargin , win_y1 + yMargin , win_x2-win_x1-xMargin-x2Margin , win_y2-win_y1-yMargin-y2Margin))

        dirI = os.path.dirname(os.path.abspath(__file__)) + "\\images\\"
        os.makedirs(dirI, exist_ok=True)

        fileName = dirI+datetime.datetime.now().strftime('%Y%m%d%H%M%S')+'.png'
        screenshot.save(fileName)

		# 3.原稿画像の読み込み
        img_org = Image.open(fileName)

		# 4.ＯＣＲ実行
        builder = pyocr.builders.TextBuilder()
        result = tool.image_to_string(img_org, lang="jpn", builder=builder)
        
        dir = os.path.dirname(os.path.abspath(__file__)) + "\\text\\"
        os.makedirs(dir, exist_ok=True)

        f = open(dir + datetime.datetime.now().strftime('%Y%m%d') +'.txt', 'a') 

        f.writelines(result) # 引数の文字列をファイルに書き込む
        f.close() # ファイルを閉じる

if __name__ == "__main__":
    app = wx.App()
    frame = MyFrame(None, wx.ID_ANY, 'ふきだし')
    icon = wx.Icon("icon.ico", wx.BITMAP_TYPE_ICO)
    frame.SetIcon(icon)

    p = wx.Panel(frame, wx.ID_ANY)
    l = wx.StaticText(p, wx.ID_ANY,'対象のプログラムを前面に出し、[Ctrl]+[Space]を押してください。\n[Alt]+[Space]で文字を解析します。' ,pos=(10,10))    

    spinctrl = wx.SpinCtrl(p, wx.ID_ANY,min=0, max=500,pos=(10,55))
    spinText = wx.StaticText(p, wx.ID_ANY,'上マージン' ,pos=(100,60))

    spinctrl2 = wx.SpinCtrl(p, wx.ID_ANY,min=0, max=500,pos=(10,85))
    spinText2 = wx.StaticText(p, wx.ID_ANY,'左マージン' ,pos=(100,90))

    spinctrl3 = wx.SpinCtrl(p, wx.ID_ANY,min=0, max=500,pos=(10,115))
    spinText3 = wx.StaticText(p, wx.ID_ANY,'右マージン' ,pos=(100,120))

    spinctrl4 = wx.SpinCtrl(p, wx.ID_ANY,min=0, max=500,pos=(10,145))
    spinText4 = wx.StaticText(p, wx.ID_ANY,'下マージン' ,pos=(100,150))

    layout = wx.BoxSizer(wx.VERTICAL)
    layout.Add(l, flag=wx.ALL, border=10)
    p.SetSizer(layout)

    frame.Show()

    app.MainLoop()