import os
import datetime

log_date = datetime.datetime.now().strftime('%Y-%m-%d')
log_path = "{}\\log\\automail_{}.log".format(os.path.dirname(__file__), log_date)

def write_log(log_info):
    log_time = datetime.datetime.now() #.strftime('%Y-%m-%d %H:%M:%S.%f')
    with open(log_path, 'a+', encoding='utf-8')as fpw:
        fpw.write("{us_time} {log_content}\n".format(us_time = log_time, log_content = log_info))