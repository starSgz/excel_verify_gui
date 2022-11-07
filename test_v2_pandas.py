import pandas as  pd
from copy import deepcopy
import datetime
data = pd.read_excel('mysqlData.xlsx')
data1 = data.append(deepcopy(data))
data2 = data1.append(deepcopy(data1))
data3 = data2.append(deepcopy(data2))
data4 = data3.append(deepcopy(data3))

print(data.head(5))

# 计算时间的装饰器
def task_content_time(func):
    start_time = datetime.datetime.now()

    def wrapper(*args, **kwargs):
        func(*args, **kwargs)
        end_time = datetime.datetime.now()
        take_time = end_time - start_time
        print(func.__name__, "任务总共花费时间:", take_time)

    return wrapper


@task_content_time
def test(mysql_verify):
    print(len(data4))
    print(any(mysql_verify['yt']=='60103'))

test(data4)