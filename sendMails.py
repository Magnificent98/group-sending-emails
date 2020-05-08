#encoding=utf-8
import wx
import os
import smtplib
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email import encoders
class MainFrame(wx.Frame):
    '''初始界面
        给出一个连接好的smtp对象
        return smtp
        return mailAddr
        return username
    '''
    def __init__(self, superior):
        wx.Frame.__init__(self, parent=superior, title=u'批量发送邮件', size=(300, 500), style=wx.CLOSE_BOX|wx.CAPTION)
        # ---------------背景设置---------------#
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour('Gray')
        wx.StaticBitmap(parent=self.panel, bitmap=wx.Image(r'guaiqiao1.png', wx.BITMAP_TYPE_ANY).ConvertToBitmap(), pos=(70, 0))
        wx.StaticBitmap(parent=self.panel, bitmap=wx.Image(r'guaiqiao2.png', wx.BITMAP_TYPE_ANY).ConvertToBitmap(), pos=(70, 370))

        #---------------选择服务器---------------#
        wx.StaticText(parent=self.panel, label='请选择服务器:', pos=(100, 110))
        self.__servers = {'QQ mail': 'smtp.qq.com:587', '163 mail': 'smtp.163.com:25', 'outlook': 'smtp.office365.com:587', 'G-mail(暂不可用)': 'smtp.gmail.com:587'}
        self.__serversCompany = {'QQ mail': 'smtp.qq.com:465', '163 mail': 'smtp.qiye.163.com:994', 'outlook': 'smtp.office365.com:587', 'G-mail(暂不可用)': 'smtp.gmail.com:465'}
        self.__suffix = {'QQ mail': '@qq.com', '163 mail': '@163.com', 'outlook': '@outlook.com', 'G-mail(暂不可用)': '@gmail.com'}
        self.__comboBox = wx.ComboBox(self.panel, value='choose here', choices=self.__servers.keys(), pos=(100, 130), size=(100, 30))
        self.Bind(wx.EVT_COMBOBOX, self.__ChooseServer, self.__comboBox)

        #---------------填写相关信息---------------#
        wx.StaticText(parent=self.panel, label='发送方的邮箱地址:', pos=(90, 170))
        self.__mailSuffix = wx.StaticText(parent=self.panel, label='', pos=(190, 190))
        self.__mailAddrGetter = wx.TextCtrl(self.panel, -1, pos=(70, 190), size=(120, 20))
        wx.StaticText(parent=self.panel, label='安全授权码:(SMTP)', pos=(90, 220))
        self.__pswdGetter = wx.TextCtrl(self.panel, -1, pos=(70, 240), size=(160, 20), style=wx.TE_PASSWORD)
        wx.StaticText(parent=self.panel, label='发送方的名字:', pos=(100, 270))
        self.__nameGetter = wx.TextCtrl(self.panel, -1, pos=(70, 290), size=(160, 20))

        # ---------------提交信息按钮---------------#
        self.__buttonSubmit = wx.Button(self.panel, -1, '连接服务器', pos=(110, 350))
        self.Bind(wx.EVT_BUTTON, self.__SubmitInfo, self.__buttonSubmit)


    def __ChooseServer(self, event):
        self.__mailSuffix.SetLabel(self.__suffix[self.__comboBox.GetValue()])
        self.server = self.__servers[self.__comboBox.GetValue()]
        self.serverCompny = self.__serversCompany[self.__comboBox.GetValue()]


    def __SubmitInfo(self, event):
        global username
        global mailAddr
        self.pswd = self.__pswdGetter.GetValue()
        username = self.__nameGetter.GetValue()
        mailAddr = self.__mailAddrGetter.GetValue()+self.__suffix[self.__comboBox.GetValue()]
        global smtp
        smtp = smtplib.SMTP()
        try:
            smtp.connect(self.server)
            smtp.ehlo()
            smtp.starttls()
        except smtplib.SMTPServerDisconnected:
            print "server disconnected"
            wx.MessageBox('端口错误！正在尝试使用企业版端口与SSL加密方式', '错误信息', wx.ICON_ERROR)
            smtp = smtplib.SMTP_SSL(self.serverCompny)
            smtp.ehlo()
        try:
            smtp.login(mailAddr, self.pswd)
            global flag
            flag = 1
            print "login successfully"
            self.Destroy()
        except smtplib.SMTPAuthenticationError:
            print "authentication error"
            wx.MessageBox('验证发生错误，请检查邮箱地址与安全授权码！', '错误信息', wx.ICON_ERROR)


class AttachFrame(wx.Frame):
    '''获取正文以及路径界面
        给出正文信息，附件绝对路径，excel表格绝对路径
        return text
        return attAddr
        return excelAddr
    '''
    def __init__(self, superior):
        wx.Frame.__init__(self, parent=superior, title='连接成功', size=(500, 300), style=wx.CLOSE_BOX|wx.CAPTION)
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour('Gray')
        # ---------------输入正文---------------#
        wx.StaticText(parent=self.panel, label='主题:', pos=(10, 10))
        self.__subjectGetter = wx.TextCtrl(self.panel, -1, pos=(10, 30), size=(200, 20))
        wx.StaticText(parent=self.panel, label='请输入正文:', pos=(10, 60))
        wx.StaticText(parent=self.panel, label='Dear XXX:', pos=(10, 80))
        self.__inputboxGetter = wx.TextCtrl(self.panel, -1, pos=(10, 110), size=(200, 150), style=wx.TE_MULTILINE|wx.VSCROLL)
        # ---------------拖拽信息---------------#
        # ---------------附件文件夹---------------#
        wx.StaticText(parent=self.panel, label='请将附件所在文件夹拖入下面的方框中:', pos=(250, 30))
        self.attPanel = wx.Panel(self.panel)
        self.attPanel.SetBackgroundColour('Yellow')
        self.attPanel.SetPosition((250, 60))
        self.attPanel.SetSize((230, 60))
        wx.StaticText(parent=self.panel, label='附件文件夹地址为:', pos=(250, 70))
        self.__attAddrLabel = wx.StaticText(parent=self.panel, label='', pos=(250, 90))
        self.__filedrop1 = FileDrop(self.__attAddrLabel)
        self.attPanel.SetDropTarget(self.__filedrop1)
        # ---------------Excel文件---------------#
        wx.StaticText(parent=self.panel, label='请将Excel表格拖入下面的方框中:', pos=(250, 140))
        self.excelPanel = wx.Panel(self.panel)
        self.excelPanel.SetBackgroundColour('Green')
        self.excelPanel.SetPosition((250, 160))
        self.excelPanel.SetSize((230, 60))
        wx.StaticText(parent=self.panel, label='Excel文件地址为:', pos=(250, 170))
        self.__excelAddrLabel = wx.StaticText(parent=self.panel, label='', pos=(250, 190))
        self.__filedrop2 = FileDrop(self.__excelAddrLabel)
        self.excelPanel.SetDropTarget(self.__filedrop2)

        # ---------------确认按钮---------------#
        self.__buttonSubmit = wx.Button(self.panel, -1, '确认', pos=(320, 240))
        self.Bind(wx.EVT_BUTTON, self.__SubmitInfo, self.__buttonSubmit)

    def __SubmitInfo(self, event):
        global subject
        global text
        global attAddr
        global excelAddr
        subject = self.__subjectGetter.GetValue()
        text = self.__inputboxGetter.GetValue()
        attAddr = self.__attAddrLabel.GetLabel()
        excelAddr = self.__excelAddrLabel.GetLabel()
        self.Destroy()


class FileDrop(wx.FileDropTarget):
    def __init__(self, AddrLabel):
        wx.FileDropTarget.__init__(self)
        self.AddrLabel = AddrLabel

    def OnDropFiles(self, x, y, filepath):
        self.AddrLabel.SetLabel(filepath[0])
        return True


class SendMails(wx.Frame):
    '''开始发送邮件

    '''
    def __init__(self, superior):
        wx.Frame.__init__(self, parent=superior, title='确认信息', size=(600, 400), style=wx.MINIMIZE_BOX|wx.CAPTION|wx.CLOSE_BOX)
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour('Gray')
        self.text = wx.TextCtrl(self.panel, -1, pos=(20, 40), size=(560, 270), style=wx.TE_MULTILINE|wx.TE_READONLY|wx.HSCROLL)
        self.__buttonSubmit = wx.Button(self.panel, -1, '开始发送', pos=(340, 340))
        self.Bind(wx.EVT_BUTTON, self.__SubmitInfo, self.__buttonSubmit)
        self.__buttonPrev = wx.Button(self.panel, -1, '上一步', pos=(200, 340))
        self.Bind(wx.EVT_BUTTON, self.__PreviousStep, self.__buttonPrev)
        self.showInfo()

    def showInfo(self):
        self.text.write("姓名\t\t邮箱地址\t\t附件名称\t\t状态\n")
        self.__sendList = self.ParseExcel()
        for items in self.__sendList:
            self.text.SetDefaultStyle(wx.TextAttr('BLACK'))
            self.text.write(items[1]+'\t\t'+items[0]+'\t\t'+items[2])
            if os.path.exists(os.path.join(attAddr, items[2])):
                self.text.SetDefaultStyle(wx.TextAttr('GREEN'))
                self.text.write('\t\tFound\n')
            else:
                self.text.SetDefaultStyle(wx.TextAttr('RED'))
                self.text.write('\t\tMISSING\n')
        self.text.SetDefaultStyle(wx.TextAttr('BLACK'))
        self.text.write('\n请确保所有的附件状态都是绿色的Found！否则不会发送。\n')
        self.text.write('\n核实信息后点击开始发送按钮。\n')

    def __PreviousStep(self, event):
        self.Destroy()
        app = wx.App()
        frame = AttachFrame(None)
        frame.Show(True)
        app.MainLoop()
        frame = SendMails(None)
        frame.Show(True)
        app.MainLoop()

    def __SubmitInfo(self, event):
        self.SetTitle('正在发送...')
        self.__buttonSubmit.Destroy()
        self.__buttonPrev.Destroy()
        self.text.Clear()
        self.packing()

    def packing(self):
        '''打包邮件，包括发件人，收件人，附件'''
        for items in self.__sendList:
            msg = MIMEMultipart("发送邮件")
            msg['Subject'] = subject
            msg['From'] = formataddr([username, mailAddr])
            msg['To'] = formataddr([items[1], items[0]])
            localtext = text
            localtext = "Dear " + items[1] + ":\n" + localtext
            msg.attach(MIMEText(localtext, 'plain', 'utf-8'))

            att = MIMEBase('application', 'octet-stream')
            att.set_payload(open(os.path.join(attAddr, items[2]), 'rb').read())
            att.add_header('Content-Disposition', 'attachment', filename=items[2])
            encoders.encode_base64(att)
            msg.attach(att)  # 添加附件
            try:
                smtp.sendmail(mailAddr, items[0], msg.as_string())
            except smtplib.SMTPDataError:
                self.SetTitle('发送失败')
                self.text.write("\n抱歉，邮箱的发件数量已经达到最大！请前往网页版验证身份信息后重试。\n")
                break
            else:
                self.text.write("发送至\t"+items[1]+" "+items[0]+" \t附件为 "+items[2]+"\t发送成功！\n")
        smtp.quit()
        self.text.write("\nComplete.\n")


    def ParseExcel(self):
        '''分析excel文件，得到用户名，邮件地址，对应的附件名'''
        '''give: recv_addr, recv_name, att_name'''
        '''返回类型是三元组列表'''
        result = []
        with open(excelAddr, 'r') as f:
            data = f.readlines()
        for items in data:
            a, b, c = items.split(' ')
            result.append((b, a, c.replace('\n','')))
        return result


if __name__ == '__main__':
    global flag
    flag = 0
    app = wx.App()
    frame = MainFrame(None)
    frame.Show(True)
    app.MainLoop()

    if flag:
        frame = AttachFrame(None)
        frame.Show(True)
        app.MainLoop()

        frame = SendMails(None)
        frame.Show(True)
        app.MainLoop()
