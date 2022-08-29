import smtplib
import email
# 负责构造文本
from email.mime.text import MIMEText
# 负责构造图片
from email.mime.image import MIMEImage
# 负责将多个对象集合起来
from email.mime.multipart import MIMEMultipart
from email.header import Header

import xlrd
import os
import datetime
import pandas as pd

from commonlib import write_log

# 全局配置
g_setting = {}

# 表头字段，对应序号
g_table_head = {}

#表格内容, key:mail key     value:row list
g_teble_content = {}

#字段定义
CONF_MAIL_HOST           = "mail_host"
CONF_MAIL_SENDER         = "mail_sender"
CONF_MAIL_PASSWORD       = "mail_password"
CONF_MAIL_CC             = "mail_cc"
CONF_MAIL_SUBJECT        = 'mail_subject'
CONF_MAIL_CONTENT_TEMPLATE = 'mail_content_template'
CONF_MAIL_APPENDIX_PATH  = "appendix_path"

CONF_XLSX_PATH           = "xlsx_path"
CONF_SHEET_NAME          = "sheet_name"
CONF_MAIL_KEY            = "mail_key"
CONF_MAIL_RECEIVER_KEY   = "mail_receiver_key"
CONF_TABLE_HEAD_ROW      = "table_head_row"


def automail_start():
    init_config()
    init_content()
    # print(g_setting)

    send_mail()

def init_config():
    work_dir = os.path.dirname(__file__)
    conf_path = "{}\\conf\\automail.ini".format(work_dir)
    global g_setting
    with open(conf_path, 'r', encoding='utf-8') as fpr:
        setting = fpr.read()
        lines = setting.split('\n')
        for line in lines:
            if line.find('=') < 0 or line.find('#') >=0:
                continue
            kvs = line.split('=')
            g_setting[kvs[0]] = kvs[1]
    write_log("Config load complete, {}".format(str(g_setting)))
    return

# 初始化表格文件内容
def init_content():
    xlsx_path = get_setting(CONF_XLSX_PATH)
    if None == xlsx_path:
        return

    sheet_name = get_setting(CONF_SHEET_NAME)
    if None == sheet_name:
        return

    mail_key_field = get_setting(CONF_MAIL_KEY)
    if None == mail_key_field:
        return

    # mail_receiver_key = get_setting(CONF_MAIL_RECEIVER_KEY)
    # if None == mail_receiver_key:
    #     return

    conf_table_head_row = get_setting(CONF_TABLE_HEAD_ROW)
    if None == conf_table_head_row:
        return
    xlsx_data = xlrd.open_workbook(xlsx_path)
    sheet_table = xlsx_data.sheet_by_name(sheet_name)
    head_row_num = int(conf_table_head_row)
    table_head = sheet_table.row_values(head_row_num)  # 获取第二行的数据
    # print("row content:", table_head)

    index = 0
    global g_table_head
    for field in table_head:
        g_table_head[field] = index
        index += 1
    # print(g_table_head)

    # sheet的行数
    rows_count = sheet_table.nrows
    # print(rows_count)

    global g_teble_content
    for row_num in range(head_row_num+1, rows_count):
        row_content = sheet_table.row_values(row_num)
        mail_key = row_content[g_table_head[mail_key_field]]
        if len(str(mail_key)) == 0:
            continue
        if mail_key in g_teble_content.keys():
            g_teble_content[mail_key].append(row_content)
        else:
            g_teble_content[mail_key] = [row_content]

    # print(g_teble_content)

def get_receivers(content_list):
    receivers_set = set()
    str_receiver_keys = get_setting(CONF_MAIL_RECEIVER_KEY)
    if None == str_receiver_keys:
        return
    receiver_keys = str_receiver_keys.split(',')

    for line in content_list:
        for key in receiver_keys:
            receivers_set.add(line[g_table_head[key]])
    # print(receivers_set)
    return receivers_set


def get_exceed_time_invoice_count(content_list):
    invoice_count = 0
    for line in content_list:
        if line[g_table_head["账龄"]] >= 90:
            invoice_count += 1
    return invoice_count



def send_mail():
    # SMTP服务器,这里使用163邮箱
    mail_host = get_setting(CONF_MAIL_HOST)
    if None == mail_host:
        return

    # 发件人邮箱
    mail_sender = get_setting(CONF_MAIL_SENDER)
    if None == mail_sender:
        return

    # 抄送人邮箱
    mail_cc = get_setting(CONF_MAIL_CC)


    # 邮件发送日期
    mail_date = datetime.datetime.now().strftime('%Y-%m-%d')



    # 邮箱授权码
    mail_license = get_setting(CONF_MAIL_PASSWORD)
    if None == mail_license:
        return

    # 邮件主题
    mail_subject = get_setting(CONF_MAIL_SUBJECT)
    if None == mail_subject:
        return

    mail_subject = mail_subject.format(config_mail_date = mail_date)
    # print("test", mail_subject)

    # 邮件正文模板
    tempate_path = get_setting(CONF_MAIL_CONTENT_TEMPLATE)
    if None == tempate_path:
        return
    mail_content_temp = ""
    with open(tempate_path, 'r', encoding='utf-8') as fpr:
        mail_content_temp = fpr.read()


    for key, val in g_teble_content.items():

        mm = MIMEMultipart('mixed')

        # 设置发送者,注意严格遵守格式,里面邮箱为发件人邮箱
        mm["From"] = generate_receiver_str(mail_sender)
        # 设置接受者,注意严格遵守格式,里面邮箱为接受者邮箱
        receivers_set = get_receivers(val)
        mm["To"] = generate_receiver_str(receivers_set)
        # print(mm["To"])

        # 设置抄送人
        if None != mail_cc:
            mm["Cc"] = mail_cc
        # 设置邮件主题
        mm["Subject"] = Header(mail_subject,'utf-8')



        # mail_content
        invoice_count = get_exceed_time_invoice_count(val)
        mail_content = mail_content_temp.format(config_mail_date = mail_date, calc_invoice_count = invoice_count)
        # print(mail_content)
        # 构造文本,参数1：正文内容，参数2：文本格式，参数3：编码方式
        # message_text = MIMEText(mail_content,"plain","utf-8")
        # 向MIMEMultipart对象中添加文本对象
        # mm.attach(message_text)

        mail_body = generate_mail_body(mail_content, val)

        # return
        context = MIMEText(mail_body, _subtype='html', _charset='utf-8')  # 解决乱码
        mm.attach(context)

        # 构造附件
        appendix_path = get_setting(CONF_MAIL_APPENDIX_PATH)
        if None != appendix_path:

            atta = MIMEText(open(appendix_path, 'rb').read(), 'base64', 'utf-8')
            # 设置附件信息
            appendix_filename = appendix_path.split('\\')[-1]
            # print(appendix_filename)
            atta["Content-Disposition"] = 'attachment; filename="{}"'.format(appendix_filename)
            # 添加附件到邮件信息当中去
            mm.attach(atta)


        # return
        # 创建SMTP对象
        stp = smtplib.SMTP()
        # 设置发件人邮箱的域名和端口，端口地址为25
        port = 25
        write_log("Try to connet mail server：{}:{}".format(mail_host, port))
        # set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
        stp.set_debuglevel(1)
        stp.connect(mail_host, port)

        # 登录邮箱，传递参数1：邮箱地址，参数2：邮箱授权码

        stp.login(mail_sender,mail_license)
        # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
        stp.sendmail(mail_sender, list(receivers_set) + mail_cc.split(','), mm.as_string())
        print("邮件发送成功，项目编码：", key)
        write_log("邮件发送成功，项目编码：{}".format(key))
        # 关闭SMTP对象
        stp.quit()



def get_setting(str_field):
    if str_field not in g_setting.keys():
        write_log("{} is not loaded".format(str_field))
        return None

    return g_setting[str_field]

def generate_receiver_str(receivers):
    if type(receivers) == str:
        return "{}<{}>".format(receivers.split('@')[0], receivers)

    res = ""
    for each in receivers:
        res += "{}<{}>,".format(each.split('@')[0], each)

    return res

def generate_mail_body(mail_text, mail_table):
    text_list = mail_text.split('\n')
    mail_text = "</br>".join(text_list)
    # print("mail text", mail_text)
    head_line = """
                <tr>
                    {}     
                </tr>"""
    units = ""
    for field in g_table_head:
        units += """<td>""" + str(field) + """</td>"""
    head_line = head_line.format(units)

    all_content_lines = ""
    for line in mail_table:
        single_content_line = """
                            <tr>
                                {}     
                            </tr>"""
        content_units = ""
        for field in line:
            content_units += """<td>""" + str(field) + """</td>"""
        all_content_lines += single_content_line.format(content_units)



    mail_body = """\
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />


    <body>
    <div id="container">
    <p style="line-height: 12px;">{html_text}</p>
    <div id="content">
     <table width="2000" border="2" bordercolor="black" cellspacing="0" cellpadding="0" >
     {html_table_head}
     {html_table_content}
    </table>
    </div>
    </div>
    </div>
    </body>
    </html>
          """.format(html_text = mail_text, html_table_head = head_line, html_table_content = all_content_lines)

    return mail_body