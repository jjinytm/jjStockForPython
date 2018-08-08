import wx
import wx.html2

class MyBrowser(wx.Dialog):
  def __init__(self, *args, **kwds):
    wx.Dialog.__init__(self, *args, **kwds)
    sizer = wx.BoxSizer(wx.VERTICAL)
    self.browser = wx.html2.WebView.New(self)
    sizer.Add(self.browser, 1, wx.EXPAND, 10)
    self.SetSizer(sizer)
    self.SetSize((700, 700))

if __name__ == '__main__':
  app = wx.App()
  dialog = MyBrowser(None, -1)
  webviewcontent = "<object width='680' height='855' id='NaverChart' classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'\
                                               codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0'><param name='movie' \
                                               value='https://ssl.pstatic.net/imgstock/fchart/NaverMashUpChart_1.0.0.swf'><param name='quality' value='high'>\
                                               <param name='FlashVars' value='Symbol=207940&amp;Description=&amp;MaxIndCount=4&amp;ChartType=캔들차트&amp;TimeFrame=day&amp;\
                                               EditMode=true&amp;DataKey=undefined&amp;ExternalInterface=false'><param name='wmode' value='opaque'><embed name='NaverChart' \
                                               width='680' height='855' id='NaverChart' pluginspage='http://www.macromedia.com/go/getflashplayer' \
                                               src='https://ssl.pstatic.net/imgstock/fchart/NaverMashUpChart_1.0.0.swf' type='application/x-shockwave-flash' \
                                               flashvars='Symbol=207940&amp;Description=&amp;MaxIndCount=4&amp;ChartType=캔들차트&amp;TimeFrame=day&amp;EditMode=true&amp;\
                                               DataKey=undefined&amp;ExternalInterface=false' wmode='opaque' quality='high' swliveconnect='TRUE'></object>"
  dialog.browser.LoadURL("https://finance.naver.com/item/fchart.nhn?code=207940")
  dialog.Show()
  app.MainLoop()