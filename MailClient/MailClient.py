import math
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase
import wx
import smtplib
import imaplib
import email
import os
import sys
import wx.richtext as rt
import wx.html2
import sqlite3

conn = sqlite3.connect('maillist.db')
c = conn.cursor()

c.execute("""CREATE TABLE IF NOT EXISTS emaillist (
            fromm text,
            subject text,
            mailcontent text,
            date text
            )""")


class MailClient(wx.Frame):
    host = 'imap.gmail.com'
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    mail = imaplib.IMAP4_SSL(host)
    your_email =''

    def __init__(self, parent, title):
        wx.Frame.__init__(self, None, title='Mail Client')

        self.InitUI()
        self.SetSize(wx.Size(500, 400))
        # self.Maximize(True)

    def InitUI(self):
        vbox = wx.BoxSizer(wx.VERTICAL)

        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        providers = ['Gmail', 'Yahoo', 'Zoho']
        cbtext = wx.StaticText(self, label="Email Provider:", style=wx.ALIGN_CENTRE)
        self.choice = wx.Choice(self, choices=providers)
        self.choice.Bind(wx.EVT_CHOICE, self.OnCombo)
        hbox3.Add(cbtext, flag=wx.LEFT, border=5)
        hbox3.Add(self.choice, flag=wx.LEFT, border=5)
        vbox.Add(hbox3, flag=wx.LEFT | wx.TOP, border=5)

        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        st1 = wx.StaticText(self, label='Username: ')
        hbox1.Add(st1, flag=wx.LEFT, border=5)
        self.emailady = wx.TextCtrl(self)
        hbox1.Add(self.emailady, proportion=0)
        vbox.Add(hbox1, flag=wx.LEFT, border=5)

        st2 = wx.StaticText(self, label='Password: ')
        hbox1.Add(st2, flag=wx.LEFT, border=5)
        self.parola = wx.TextCtrl(self, style=wx.TE_PASSWORD)
        hbox1.Add(self.parola, proportion=0)
        vbox.Add(hbox1, flag=wx.LEFT, border=5)

        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        loginButton = wx.Button(self, label='Login')
        hbox2.Add(loginButton, flag=wx.LEFT, border=90)
        vbox.Add(hbox2, flag=wx.LEFT | wx.TOP, border=0)
        loginButton.Bind(wx.EVT_BUTTON, self.Login)

        closeButton = wx.Button(self, label='Close')
        hbox2.Add(closeButton, flag=wx.LEFT, border=20)
        vbox.Add(hbox2, flag=wx.LEFT | wx.TOP, border=0)
        closeButton.Bind(wx.EVT_BUTTON, self.onClose)

        self.png = wx.StaticBitmap(self,-1,wx.Bitmap("logomail.png", wx.BITMAP_TYPE_ANY))
        vbox.Add(self.png,flag = wx.ALL,border=5)
        self.SetSizer(vbox)

    def OnCombo(self, e):
        prov = self.choice.GetSelection()
        print(prov)
        if prov == 2:
            MailClient.host = 'imap.zoho.com'
            MailClient.server = smtplib.SMTP_SSL('smtp.zoho.com', 465)
        elif prov == 1:
            MailClient.host = 'pop.mail.yahoo.com'
            MailClient.server = smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465)
        elif prov == 0:
            MailClient.host = 'imap.gmail.com'
            MailClient.server = smtplib.SMTP_SSL('smtp.gmail.com', 465)

    def Login(self, e):
        MailClient.your_email = self.emailady.GetValue()
        your_password = self.parola.GetValue()

        MailClient.server.ehlo()
        MailClient.server.login(MailClient.your_email, your_password)
        MailClient.mail.login(MailClient.your_email, your_password)

        secondWindow = Inbox(None, 'Inbox')
        self.Hide()
        secondWindow.Show()

    def onClose(self, e):
        MailClient.server.close()
        conn.close()
        self.Close()


class Inbox(MailClient):
    subjectdir = ''
    datedir = ''
    fromdir = ''

    def __init__(self, parent, title):
        wx.Frame.__init__(self, None, title='Mail Client')
        self.Maximize(True)

        vbox = wx.BoxSizer(wx.HORIZONTAL)
        self.SetSizer(vbox)
        gs = wx.GridSizer(3, 0, 5, 5)

        gs.AddMany([(wx.Button(self, id=1, label='Inbox'), 0, 0, 0),
                    (wx.Button(self, id=2, label='New Email'), 0, 0, 0),
                    (wx.Button(self, id=3, label='Exit'), 0, 0, 0), ])

        self.Bind(wx.EVT_BUTTON, self.onnClose)
        self.Bind(wx.EVT_BUTTON, self.newEmail)
        vbox.Add(gs, proportion=0, flag=wx.LEFT | wx.TOP, border=5)
        vbox.Add(wx.StaticLine(self, -1, style=wx.LI_VERTICAL), 0, wx.EXPAND | wx.LEFT, 5)

        width = self.GetClientSize().width / 3 - 25

        self.list_ctrl = wx.ListCtrl(self, id=10, style=wx.LC_REPORT | wx.BORDER_SUNKEN)
        self.list_ctrl.InsertColumn(0, 'From', width=int(width))
        self.list_ctrl.InsertColumn(1, 'Subject', width=int(width))
        self.list_ctrl.InsertColumn(2, 'Date', width=int(width))

        vbox.Add(self.list_ctrl, 0, wx.EXPAND | wx.ALL, 5)

        self.readEmails()

        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.openEmail, self.list_ctrl)
        self.Bind(wx.EVT_BUTTON, self.refreshMail)

    def newEmail(self, e):
        e.Skip()
        if 2 == e.GetId():
            newmailwindow = SendNewMail(None, 'Send New Email')
            newmailwindow.Show()

    def refreshMail(self, e):
        e.Skip()
        if 1 == e.GetId():
            self.readEmails()

    def openEmail(self, e):
        currentItem = e.GetIndex()
        Inbox.fromdir = self.list_ctrl.GetItemText(currentItem, 0)
        Inbox.subjectdir = self.list_ctrl.GetItemText(currentItem, 1)
        Inbox.datedir = self.list_ctrl.GetItemText(currentItem, 2)

        c.execute("SELECT mailcontent FROM emaillist WHERE date = ?", (Inbox.datedir,))
        emailcontent = c.fetchone()
        emaildir = emailcontent[0]
        # subjdir = Inbox.subjectdir.replace(":", "-")
        # datdir = Inbox.datedir.replace(":", "-")
        # cwd = os.getcwd()
        # emaildir = "file:///" + cwd + "/emails/" + datdir + "/" + subjdir + "/msg-part-00000002.html"
        thirdWindow = OpenMail(None, 'Email')
        thirdWindow.Show()
        thirdWindow.openEmail2(emaildir)

    def readEmails(self):
        MailClient.mail.select("inbox")
        result, data = MailClient.mail.uid('search', None, "ALL")

        inbox_item_list = data[0].split()
        index = 0
        inbox_item_list.reverse()

        for item in inbox_item_list[0:5]:
            result2, email_data = MailClient.mail.uid('fetch', item, '(RFC822)')
            raw_email = email_data[0][1].decode("cp1252", errors='ignore')
            email_message = email.message_from_string(raw_email)
            to_ = email_message['To']
            from_ = email_message['From']
            subject_ = email_message['Subject']
            date_ = email_message['date']
            counter = 1
            for part in email_message.walk():
                mcontext = part.get_payload(decode=True)
                counter += 1
            c.execute("INSERT INTO emaillist VALUES (?, ?, ?, ?)", (from_, subject_, mcontext, date_))
            # save_path = os.path.join(os.getcwd(), "emails", date_.replace(":", "-"), subject_.replace(":", "-"))
            # if not os.path.exists(save_path):
            #    os.makedirs(save_path)
            # with open(os.path.join(save_path, filename), 'wb') as fp:
            #    fp.write(part.get_payload(decode=True))

            self.list_ctrl.InsertItem(index, from_)
            self.list_ctrl.SetItem(index, 1, subject_)
            self.list_ctrl.SetItem(index, 2, date_)

            index += 1

    def onnClose(self, e):
        e.Skip()
        if 3 == e.GetId():
            MailClient.server.close()
            self.Close()
            x = MailClient(None, 'MailClient')
            x.onClose(e)


class OpenMail(Inbox):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, None, title='Email')

        vbox = wx.BoxSizer(wx.VERTICAL)

        fgs = wx.FlexGridSizer(3, 1, 5, 0)
        fromdir = wx.TextCtrl(self, value='From: ' + Inbox.fromdir, style=wx.TE_READONLY | wx.EXPAND)
        datedir = wx.TextCtrl(self, value='Date: ' + Inbox.datedir, style=wx.TE_READONLY | wx.EXPAND)
        subjdir = wx.TextCtrl(self, value='Subject: ' + Inbox.subjectdir, style=wx.TE_READONLY | wx.EXPAND)

        fgs.AddMany([(fromdir, 1, wx.EXPAND),
                     (datedir, 1, wx.EXPAND),
                     (subjdir, 1, wx.EXPAND), ])

        fgs.AddGrowableCol(0, 1)

        vbox.Add(fgs, proportion=0, flag=wx.ALL | wx.EXPAND, border=5)

        self.htmlwin = wx.html2.WebView.New(self)
        vbox.Add(self.htmlwin, 1, wx.EXPAND, 10)
        self.SetSizer(vbox)
        self.SetSize(600, 600)

    def openEmail2(self, x):
        self.htmlwin.SetPage(x, "")


class SendNewMail(Inbox):
    message2 = ''
    filepath = ''

    def __init__(self, parent, title):
        wx.Frame.__init__(self, None, title='Send New Email')
        self.filepaths = []
        self.currentDir = os.path.abspath(os.path.dirname(sys.argv[0]))
        self.SendMailUI()
        self.SetSize(700, 700)


    def SendMailUI(self):

        panel = wx.Panel(self)
        # ~ """
        # ~ Текстовите полета със разделители в първа колона (от,до,тема,текст)
        # ~ и полетата за писане във втората колона(отПоле,доПоле,тема,текст)
        # ~ направени със флексгрид вкаран във вертикален box sizer
        # ~ """
        verticalBox = wx.BoxSizer(wx.VERTICAL)

        textFields = wx.FlexGridSizer(6, 2, 9, 25)

        ot = wx.StaticText(panel, label="От :")
        to = wx.StaticText(panel, label="До :")
        subject = wx.StaticText(panel, label="Тема :")
        attachedFiles = wx.StaticText(panel, label="Прикачени файлове :")

        otField = wx.TextCtrl(panel)
        self.toField = wx.TextCtrl(panel)
        self.subjectField = wx.TextCtrl(panel)
        self.textField = rt.RichTextCtrl(panel, style=wx.TE_MULTILINE)
        self.attachTxt = wx.TextCtrl(panel, wx.ID_ANY, '', style=wx.TE_MULTILINE)
        # ~ self.attachTxt.Disable()

        # ~ buttonPanel=wx.Panel(panel)
        # ~ buttonPanel.SetBackgroundColour((0,0,0))

        # ~ """
        # ~ Custom Бутони за форматиране
        # ~ във хоризонтален box sizer
        # ~ """

        vToolBar = wx.BoxSizer(wx.VERTICAL)

        bmpBold = wx.Bitmap("bold.png", wx.BITMAP_TYPE_PNG)
        bmpButtonBold = wx.BitmapButton(panel, id=wx.ID_ANY, bitmap=bmpBold,
                                        size=(bmpBold.GetWidth() + 10, bmpBold.GetHeight() + 10))
        bmpButtonBold.Bind(wx.EVT_BUTTON, self.OnBold)
        vToolBar.Add(bmpButtonBold, 0, wx.ALIGN_LEFT)

        bmpItalic = wx.Bitmap("Italicbutton.png", wx.BITMAP_TYPE_PNG)
        bmpButtonItalic = wx.BitmapButton(panel, id=wx.ID_ANY, bitmap=bmpItalic,
                                          size=(bmpItalic.GetWidth() + 10, bmpItalic.GetHeight() + 10))
        bmpButtonItalic.Bind(wx.EVT_BUTTON, self.OnItalic)
        vToolBar.Add(bmpButtonItalic, 0, wx.ALIGN_LEFT)

        bmpUnderline = wx.Bitmap("Underlineicon.png", wx.BITMAP_TYPE_PNG)
        bmpButtonUnderline = wx.BitmapButton(panel, id=wx.ID_ANY, bitmap=bmpUnderline,
                                             size=(bmpUnderline.GetWidth() + 10, bmpUnderline.GetHeight() + 10))
        bmpButtonUnderline.Bind(wx.EVT_BUTTON, self.OnUnderline)
        vToolBar.Add(bmpButtonUnderline, 0, wx.ALIGN_LEFT)

        bmpAlignLeft = wx.Bitmap("AlignLeft.png", wx.BITMAP_TYPE_PNG)
        bmpButtonAlignLeft = wx.BitmapButton(panel, id=wx.ID_ANY, bitmap=bmpAlignLeft,
                                             size=(bmpAlignLeft.GetWidth() + 10, bmpAlignLeft.GetHeight() + 10))
        bmpButtonAlignLeft.Bind(wx.EVT_BUTTON, self.OnAlignLeft)
        vToolBar.Add(bmpButtonAlignLeft, 0, wx.ALIGN_LEFT)

        bmpAlignCenter = wx.Bitmap("AlignCenter.png", wx.BITMAP_TYPE_PNG)
        bmpButtonAlignCenter = wx.BitmapButton(panel, id=wx.ID_ANY, bitmap=bmpAlignCenter,
                                               size=(bmpAlignCenter.GetWidth() + 10, bmpAlignCenter.GetHeight() + 10))
        bmpButtonAlignCenter.Bind(wx.EVT_BUTTON, self.OnAlignCenter)
        vToolBar.Add(bmpButtonAlignCenter, 0, wx.ALIGN_LEFT)

        bmpAlignRight = wx.Bitmap("AlignRight.png", wx.BITMAP_TYPE_PNG)
        bmpButtonAlignRight = wx.BitmapButton(panel, id=wx.ID_ANY, bitmap=bmpAlignRight,
                                              size=(bmpAlignRight.GetWidth() + 10, bmpAlignRight.GetHeight() + 10))
        bmpButtonAlignRight.Bind(wx.EVT_BUTTON, self.OnAlignRight)
        vToolBar.Add(bmpButtonAlignRight, 0, wx.ALIGN_LEFT)

        bmpAttach = wx.Bitmap("Attach.png", wx.BITMAP_TYPE_PNG)
        bmpButtonAttach = wx.BitmapButton(panel, id=wx.ID_ANY, bitmap=bmpAttach,
                                          size=(bmpAttach.GetWidth() + 10, bmpAttach.GetHeight() + 10))
        bmpButtonAttach.Bind(wx.EVT_BUTTON, self.OnAttach)
        vToolBar.Add(bmpButtonAttach, 0, wx.ALIGN_LEFT)

        # ~ buttonPanel.SetSizer(hToolBar)

        # ~ """
        # ~ Вкарване на полетата във flexgrid sizser-a
        # ~ """

        textFields.AddMany(
            [(ot), (otField, 1, wx.EXPAND), (to), (self.toField, 1, wx.EXPAND), (subject), (self.subjectField, 1, wx.EXPAND),
             (vToolBar),
             (self.textField, 1, wx.EXPAND), (attachedFiles), (self.attachTxt, 1, wx.EXPAND)])

        textFields.AddGrowableRow(3, 1)
        textFields.AddGrowableCol(1, 1)

        verticalBox.Add(textFields, proportion=1, flag=wx.ALL | wx.EXPAND, border=15)

        # ~ """
        # ~ Send и Cancel бутони във вертикален box sizer(hButtonBox)
        # ~ вкарани във главния вертикален(verticalBox)
        # ~ sizer подравнени в ляво
        # ~ """

        hButtonBox = wx.BoxSizer(wx.HORIZONTAL)
        btn1 = wx.Button(panel, label='Send', size=(70, 30))
        btn1.Bind(wx.EVT_BUTTON, self.sendMail)
        hButtonBox.Add(btn1)
        btn2 = wx.Button(panel, label='Cancel', size=(70, 30))
        btn2.Bind(wx.EVT_BUTTON, self.OnClose)
        hButtonBox.Add(btn2, flag=wx.LEFT | wx.BOTTOM, border=5)

        verticalBox.Add(hButtonBox, flag=wx.ALIGN_RIGHT | wx.RIGHT, border=14)

        # ~ verticalBox.Add(hToolBar,flag=wx.ALIGN_CENTER)
        panel.SetSizer(verticalBox)

        self.textField.Bind(rt.EVT_RICHTEXT_CHARACTER, self.msgstr)
    # ~ """
    # ~ Функции за бутоните
    # ~ """

    def msgstr(self,e):
        SendNewMail.message2 = self.textField.GetValue()

    def sendMail(self, e):

        msg = MIMEMultipart('alternative')
        msg['Subject'] = self.subjectField.GetValue()
        msg['From'] = MailClient.your_email
        msg['To'] = self.toField.GetValue()

        #if SendNewMail.message2 is None:
           # SendNewMail.message2 = self.textField.GetValue()

        htmlmail = MIMEText(SendNewMail.message2, 'html')
        msg.attach(htmlmail)

        if SendNewMail.filepath != '':
            filename = os.path.basename(SendNewMail.filepath)
            attachment = open(SendNewMail.filepath, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            "attachment; filename= %s" % filename)
            msg.attach(part)

        try:
            MailClient.server.sendmail(MailClient.your_email, self.toField.GetValue(), msg.as_string())
            print('Email to {} successfully sent!\n\n'.format(self.toField.GetValue()))
        except Exception as e:
            print('Email to {} could not be sent :( because {}\n\n'.format(self.toField.GetValue(), str(e)))

        self.Close()

    def OnClose(self, e):
        self.Close(True)

    def OnBold(self, e):
        self.textField.ApplyBoldToSelection()
        SendNewMail.message2 = SendNewMail.message2.replace(self.textField.GetStringSelection(),
                                                          '<b>'+self.textField.GetStringSelection()+'</b>')

    def OnItalic(self, e):
        self.textField.ApplyItalicToSelection()
        SendNewMail.message2 = SendNewMail.message2.replace(self.textField.GetStringSelection(),
                                                          '<i>' + self.textField.GetStringSelection() + '</i>')

    def OnUnderline(self, e):
        self.textField.ApplyUnderlineToSelection()
        SendNewMail.message2 = SendNewMail.message2.replace(self.textField.GetStringSelection(),
                                                          '<u>' + self.textField.GetStringSelection() + '</u>')

    def OnAlignLeft(self, e):
        self.textField.ApplyAlignmentToSelection(wx.TEXT_ALIGNMENT_LEFT)
        SendNewMail.message2 = SendNewMail.message2.replace(self.textField.GetStringSelection(),
                                                           '<div align="left">' + self.textField.GetStringSelection() + '</div>')

    def OnAlignCenter(self, e):
        self.textField.ApplyAlignmentToSelection(wx.TEXT_ALIGNMENT_CENTRE)
        SendNewMail.message2 = SendNewMail.message2.replace(self.textField.GetStringSelection(),
                                                           '<div align="center">' + self.textField.GetStringSelection() + '</div>')

    def OnAlignRight(self, e):
        self.textField.ApplyAlignmentToSelection(wx.TEXT_ALIGNMENT_RIGHT)
        SendNewMail.message2 = SendNewMail.message2.replace(self.textField.GetStringSelection(),
                                                           '<div align="right">' + self.textField.GetStringSelection() + '</div>')

    def OnAttach(self, event):
        # ~ '''
        # ~ Displays a File Dialog to allow the user to choose a file
        # ~ and then attach it to the email.
        # ~ '''
        attachments = self.attachTxt.GetLabel()
        # create a file dialog
        wildcard = "All files (*.*)|*.*"
        dialog = wx.FileDialog(None, 'Choose a file', self.currentDir, '', wildcard, wx.FD_OPEN)
        # if the user presses OK, get the path
        if dialog.ShowModal() == wx.ID_OK:
            self.attachTxt.Show()

            SendNewMail.filepath = dialog.GetPath()
            # Change the current directory to reflect the last dir opened
            os.chdir(os.path.dirname(SendNewMail.filepath))
            self.currentDir = os.getcwd()
            # add the user's file to the filepath list
            if SendNewMail.filepath != '':
                self.filepaths.append(SendNewMail.filepath)
            # get file size
            fSize = self.getFileSize(SendNewMail.filepath)
            # modify the attachment's label based on it's current contents
            if attachments == '':
                attachments = '%s (%s)' % (os.path.basename(SendNewMail.filepath), fSize)
            else:
                temp = '%s (%s)' % (os.path.basename(SendNewMail.filepath), fSize)
                attachments = attachments + '; ' + temp
            self.attachTxt.SetLabel(attachments)
        dialog.Destroy()

    def getFileSize(self, f):

        fSize = os.stat(f).st_size
        if fSize >= 1073741824:  # gigabyte
            fSize = int(math.ceil(fSize / 1073741824.0))
            size = '%s GB' % fSize
        elif fSize >= 1048576:  # megabyte
            fSize = int(math.ceil(fSize / 1048576.0))
            size = '%s MB' % fSize
        elif fSize >= 1024:  # kilobyte
            fSize = int(math.ceil(fSize / 1024.0))
            size = '%s KB' % fSize
        else:
            size = '%s bytes' % fSize
        return size


def main():
    app = wx.App()
    ex = MailClient(None, title='MailClient')
    ex.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
