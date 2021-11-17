"""
============================
author:qideping
time:2021/11/16 2:08 下午
E-mail:qideping@makenv.com
============================
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart  # 附件
from email.header import Header


class TestMail:
    def __init__(self):
        # 第三方 SMTP 服务
        self.mail_host = "smtp.qq.com"  # 设置服务器
        self.mail_user = "874521869@qq.com"  # 用户名
        self.mail_pass = "hmwsluvudyubbdeh"  # 口令,口令为smtp的授权码，不是账号密码

    def send_mail(self):
        sender = '874521869@qq.com'   # 发件人邮箱
        receivers = ['qideping@makenv.com']   # 收件人邮箱
        # 发送html形式的测试报告
        # file_path = "/Users/qideping/my_project/auto_test_api/app/reports/2021-11-09 10.45.49测试报告.html"
        # file = open(file_path)
        # mail_body = file.read()
        # message = MIMEText(mail_body, 'html', 'utf-8')  # 邮件内容
        message = MIMEMultipart()
        message.attach(MIMEText('Dear all:\n接口自动化测试报告详情见附件\n', 'plain', 'utf-8'))   # 邮件内容
        message['From'] = Header("测试组", 'utf-8')
        message['To'] = Header("测试负责人", 'utf-8')
        subject = '接口自动化测试报告'   # 邮件主题
        message['Subject'] = Header(subject)

        # 构造附件
        file_path = "/Users/qideping/my_project/auto_test_api/app/reports/2021-11-09 10.45.49测试报告.html"
        # file_path = "/Users/qideping/my_project/makenv/git_exercise/test.txt"
        file = open(file_path, 'rb').read()
        att = MIMEText(file, 'html', 'utf-8')
        # att = MIMEText(open("/Users/qideping/my_project/makenv/git_exercise/test.txt", "rb").read(), 'base64', 'utf-8')

        att["Content-Type"] = 'application/octet-stream'
        att["Content-Disposition"] = 'attachment; filename="测试报告.html"'
        message.attach(att)

        try:
            smtp = smtplib.SMTP()
            smtp.connect(self.mail_host, 25)  # 25 为 SMTP 端口号
            smtp.login(self.mail_user, self.mail_pass)
            smtp.sendmail(sender, receivers, message.as_string())
            print("邮件发送成功")
            smtp.quit()
        except smtplib.SMTPException as e:
            print("Error: 无法发送邮件，{}".format(e))


if __name__ == '__main__':
    send = TestMail()
    send.send_mail()



