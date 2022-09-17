# autoMail
auto mail tools. author:BrookTaylor

automail_release gui ver 1.0 2022.09.18

确保表格符合以下要求
1.首行为表头，不要有重复字段
2.日期需要转化成文本，使用公式  =TEXT(unit_pos, "YYYY-MM-DD"),否则邮件中日期可能异常
3.整型数需要转化为文本，使用公式 =TEXT(unit_pos, 0),否则邮件中日期可能异常


配置好配置文件
conf/aotomail.ini


邮件内容模板
conf/mail_template.txt