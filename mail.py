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
import traceback

from commonlib import write_log

# 全局配置
g_setting = {}

# 表头字段，对应序号
g_table_head = {}

#表格内容, key:mail key     value:row list
g_teble_content = {}

#邮件字典,key id   value CMailObj
g_mail_dict = {}

class CMailObj:
    def __init__(self, id=0, sender='', receiver=set(), cc='', subject = '', appendix = [], content = ''):
        self.id = id                  # 邮件ID
        self.sender     = sender      # 发件人
        self.receiver   = receiver    # 收件人
        self.cc         = cc          # 抄送人
        self.subject    = subject     # 邮件主题
        self.appendix   = appendix    # 附件信息
        self.content    = content     # 邮件内容

    def show(self):
        print_content = "id:{}\nsender:{}\nreceiver:{}\ncc:{}\nsubject:{}\nappendix:{}\ncontent:{}\n".format(self.id,
                                                                                                             self.sender,
                                                                                                             self.receiver,
                                                                                                             self.cc,
                                                                                                             self.subject,
                                                                                                             self.appendix,
                                                                                                             self.content)
        print(print_content)

class CMailCommon:
    def __init__(self, host='', port=0, sender='', password='', mail_date='', sended_record_path = '', sended_set=set(), selected_set=set()):
        self.host = host
        self.port = port
        self.sender = sender
        self.password = password
        self.mail_date = mail_date
        self.sended_record_path = sended_record_path
        self.sended_set = sended_set
        self.selected_set = selected_set

    def init_sended_set(self):
        if os.path.exists(self.sended_record_path):
            with open(self.sended_record_path, 'r') as fp:
                record_list = fp.read().split("\n")
                self.sended_set = set(record_list)

    def dump_sended_set(self):
        with open(self.sended_record_path, 'w') as fp:
            fp.write('\n'.join(self.sended_set))

    def clear_sended_set(self):
        if len(self.sended_set) == 0:
            return
        self.sended_set = set()
        self.dump_sended_set()

    def add_sended_set(self, mail_id):
        self.sended_set.add(mail_id)

    def remove_sended_set(self, mail_id):
        if mail_id not in self.sended_set:
            return
        self.sended_set.remove(mail_id)

    def is_in_sended_set(self, mail_id):
        return mail_id in self.sended_set

    def add_selected_mail(self, id):
        self.selected_set.add(id)

    def remove_selected_mail(self, id):
        self.selected_set.remove(id)

    def is_selected_joint_sended(self):
        # print(self.sended_set)
        # print(self.selected_set)
        # print(self.sended_set.isdisjoint(self.selected_set))
        return not (self.sended_set.isdisjoint(self.selected_set))

    def send_selected_mail(self):
        for mail_id in self.selected_set:

            if mail_id not in g_mail_dict.keys():
                write_log("Mial id not Exist in mail dict, mail_id={}".format(mail_id))
                continue
            mail_obj = g_mail_dict[mail_id]
            send_mail(mail_obj)
            # print("邮件发送成功，项目编码:", mail_id)
            self.add_sended_set(mail_id)
        self.dump_sended_set()

# 全局邮件公共信息
g_mail_common = CMailCommon()

#字段定义
CONF_MAIL_HOST           = "mail_host"
CONF_MAIL_PORT           = "port"
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
CONF_SENDED_RECORD_PATH  = "sended_record_path"


def automail_start():
    init_config()
    init_content()
    # print(g_setting)

    init_mail_dict()

    # send_all_mail()

def init_config():
    conf_path = r"conf\automail.ini"
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
    table_head = sheet_table.row_values(head_row_num)
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

# 获取账龄告警数量
def get_exceed_time_invoice_count(content_list):
    invoice_count = 0
    for line in content_list:
        if line[g_table_head["账龄"]] >= 90:
            invoice_count += 1
    return invoice_count

# 初始化邮件内容字典
def init_mail_dict():
    global g_mail_common

    # SMTP服务器,这里使用163邮箱
    mail_host = get_setting(CONF_MAIL_HOST)
    if None == mail_host:
        return
    g_mail_common.host = mail_host

    # 邮件服务器端口
    mail_port = get_setting(CONF_MAIL_PORT)
    if None == mail_port:
        return
    g_mail_common.port = mail_port

    # 发件人邮箱
    mail_sender = get_setting(CONF_MAIL_SENDER)
    if None == mail_sender:
        return
    g_mail_common.sender = mail_sender

    # 邮箱授权码
    mail_license = get_setting(CONF_MAIL_PASSWORD)
    if None == mail_license:
        return
    g_mail_common.password = mail_license

    # 邮件发送日期
    mail_date = datetime.datetime.now().strftime('%Y-%m-%d')
    g_mail_common.mail_date = mail_date

    # 获取已发送邮件记录
    sended_record_path = get_setting(CONF_SENDED_RECORD_PATH)
    if None == sended_record_path:
        return
    g_mail_common.sended_record_path = sended_record_path
    g_mail_common.init_sended_set()

    # 抄送人邮箱
    mail_cc = get_setting(CONF_MAIL_CC)

    # 邮件主题
    mail_subject = get_setting(CONF_MAIL_SUBJECT)
    if None == mail_subject:
        return

    mail_subject = mail_subject.format(config_mail_date=mail_date)
    # print("test", mail_subject)

    # 邮件附件列表
    appendix_list = []
    appendix_path = get_setting(CONF_MAIL_APPENDIX_PATH)
    if None != appendix_path:
        appendix_list = appendix_path.split(';')

    # 邮件正文模板
    tempate_path = get_setting(CONF_MAIL_CONTENT_TEMPLATE)
    if None == tempate_path:
        return
    mail_content_temp = ""
    with open(tempate_path, 'r', encoding='utf-8') as fpr:
        mail_content_temp = fpr.read()

    # 初始化邮件
    global g_mail_dict
    for key, val in g_teble_content.items():
        # 收件人集合
        receivers_set = get_receivers(val)

        # 邮件内容初始化
        invoice_count = get_exceed_time_invoice_count(val)
        mail_text = mail_content_temp.format(config_mail_date = mail_date, calc_invoice_count = invoice_count)
        mail_content = generate_mail_body(mail_text, val)

        g_mail_dict[key] = CMailObj(id = key,
                                    sender = mail_sender,
                                    receiver = receivers_set,
                                    cc = mail_cc,
                                    subject = mail_subject,
                                    appendix = appendix_list,
                                    content = mail_content)
        # g_mail_dict[key].show()
    return


def send_all_mail():
    global g_mail_dict
    for key, mail_obj in  g_mail_dict.items():
        send_mail(mail_obj)

def send_mail(mail_obj):
    print("start send mail")
    global g_mail_common
    # 创建SMTP对象设置发件人邮箱的域名和端口，端口地址
    stp = smtplib.SMTP(g_mail_common.host, g_mail_common.port)
    stp.ehlo()
    stp.starttls()

    # 登录邮箱，传递参数1：邮箱地址，参数2：邮箱授权码
    write_log("Try to connect mail server：{}:{}".format(g_mail_common.host, g_mail_common.port))

    try:
        stp.login(g_mail_common.sender, g_mail_common.password)
    except Exception as e:
        write_log("Connect mail server failed, error info: {}".format(traceback.format_exc()))

    write_log("Connect mail server succeed")


    mm = MIMEMultipart('mixed')

    # 设置发送者,注意严格遵守格式,里面邮箱为发件人邮箱
    mm["From"] = generate_receiver_str(g_mail_common.sender)

    # 设置接受者,注意严格遵守格式,里面邮箱为接受者邮箱
    mm["To"] = generate_receiver_str(mail_obj.receiver)


    # 设置抄送人
    mm["Cc"] = mail_obj.cc

    # 设置邮件主题
    mm["Subject"] = Header(mail_obj.subject, 'utf-8')

    # 邮件内容
    context = MIMEText(mail_obj.content, _subtype='html', _charset='utf-8')  # 解决乱码
    mm.attach(context)

    # 添加附件
    for each_appendix in mail_obj.appendix:
        if len(each_appendix) == 0:
            continue

        atta = MIMEText(open(each_appendix, 'rb').read(), 'base64', 'utf-8')

        # 设置附件信息
        appendix_filename = each_appendix.split('\\')[-1]
        atta['Content-Type'] = 'application/octet-stream'
        atta.add_header("Content-Disposition", 'attachment', filename = ('gbk', '', appendix_filename))

        # 添加附件到邮件信息当中去
        mm.attach(atta)

    # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
    stp.sendmail(g_mail_common.sender, list(mail_obj.receiver) + mail_obj.cc.split(';'), mm.as_string())
    print("邮件发送成功，项目编码：", mail_obj.id)
    write_log("邮件发送成功，项目编码：{}".format(mail_obj.id))

    # 关闭SMTP对象
    stp.quit()


# 获取配置字段
def get_setting(str_field):
    if str_field not in g_setting.keys():
        write_log("{} is not loaded".format(str_field))
        return None

    return g_setting[str_field]

# 生成邮件收件人 返回格式: receiver_name1<receiver_mail1>,receiver_name2<receiver_mail2>,
def generate_receiver_str(receivers):
    if type(receivers) == str:
        return "{}<{}>".format(receivers.split('@')[0], receivers)

    res = ""
    for each in receivers:
        res += "{}<{}>,".format(each.split('@')[0], each)

    return res


# 生成html邮件正文,需要生成表格,故采用html格式的邮件正文
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
    <p style="line-height: 20px;">{html_text}</p>
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