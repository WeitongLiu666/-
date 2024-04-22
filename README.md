import shutil
import time
import os
import pandas as pd
import datetime
import requests
import json
import selenium
from Crypto.Cipher import DES
from Crypto.Util.Padding import unpad
from Crypto.Hash import MD5
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import xlwings as xw
from loguru import logger
from binascii import unhexlify
from pandas_market_calendars import get_calendar
########################################################################################################################
update_excel_path = r"D:\SynologyDrive\基金代码&名称对照表.xlsx"                                                             # 基金净值表地址
user_data_dir = r"C:\\Users\ediso\AppData\Local\Google\Chrome\\User Data\Default"                                        # 替换为浏览器输入chrome://version/里的个人资料路径
value_path = r"D:\SynologyDrive\Data_Base\Invested_Fund_Value"
download_path = r"C:\Users\ediso\Documents\Downloads"                                                                   # 下载路径
huofuniu_username = "15940859777"
huofuniu_password = "Aa666888"
########################################################################################################################
start_time = time.time()
driver_service = Service(ChromeDriverManager().install())                                                               # 配置 Chrome WebDriver 服务
chrome_options = Options()                                                                                              # 配置 Chrome WebDriver 选项
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")                                                         # 配置 Chrome WebDriver 个人资料库
driver = webdriver.Chrome(service=driver_service, options=chrome_options)                                               # 创建 Chrome WebDriver 实例
url = "https://mp.fof99.com/fund/all"                                                                                   # 打开指定页面
fund_list = ['银尊1', '鑫c1', '嘉行1', '日鑫1', '嘉行2', '日鑫2', '科技兴国1', '嘉研1', '嘉研3', '黄金增强1', '尊量1', '因聚', '申睿8',
             '于航', '嘉研2', '鑫c6', '鑫光1', '申睿11', '毅多1', '量星11', '领恒', '昊倍1', '泰君26', '冲资1', '宝璟1', '田1106',
             '锡添', '山和', '嘉研倍', '嘉如', '中牛', '宝进', '银善6']

# 登录火富牛
def get_huofuniu_login(huofuniu_username, huofuniu_password):
    driver.get(url)
    time.sleep(5)
    driver.get(url)
    try:
        driver.find_element(By.XPATH,'//*[@id="username"]').clear()                                               # 清除历史登录信息
        driver.find_element(By.XPATH,'//*[@id="username"]').send_keys(huofuniu_username)                          # 填写用户名
        driver.find_element(By.XPATH,'//*[@id="password"]').clear()                                               # 清除历史登录信息
        driver.find_element(By.XPATH,'//*[@id="password"]').send_keys(huofuniu_password)                          # 填写密码
        driver.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/form/div[4]/div/div/span/button').click()  # 登录
        print("火富牛登录成功")
    except:
        pass
    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="fof-layout"]/header/div[1]/div[2]/div[1]/span[7]/a/span').click()     # 点击运维
    time.sleep(1)
    driver.find_element(By.XPATH,'//*[@id="fof-layout"]/div[1]/div/div[2]/div/div[1]/div/div/div[1]/div[1]/span[3]').click()  # 点击内部
def update_value(count):
    fundname = driver.find_element(By.XPATH,'//*[@id="fof-layout"]/div[1]/div/div[2]/div/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr[%s]/td[3]/span' % count).get_attribute("innerText")  # 获取基金名称
    fund_index = df_update[df_update['产品简称'] == fundname].index[0]                                                    # 获取产品在净值表里的index
    try:
        value_date = df_update['净值日期'][fund_index].strftime('%Y-%m-%d')                                               # 获取基金最新净值日期
        if count > 1:           # 滚动窗口
            driver.execute_script("arguments[0].scrollIntoView();", driver.find_element(By.XPATH,'//*[@id="fof-layout"]/div[1]/div/div[2]/div/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr[%s]/td[3]/span' % (count - 1)))
        time.sleep(1)
        driver.find_element(By.XPATH,'//*[@id="fof-layout"]/div[1]/div/div[2]/div/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody/tr[%s]/td[8]/a/a/i' % count).click()  # 点击净值表
        time.sleep(1)
        driver.switch_to.window(driver.window_handles[-1])                                                              # 切换到新打开的页面
        time.sleep(1)
        fundname_test = driver.find_element(By.XPATH,'//*[@id="fof-layout"]/div[1]/div/div[2]/div/div[1]/h2').get_attribute("innerText")  # 获取基金名称
        if fundname_test == fundname:
            try:
                web_date = driver.find_element(By.XPATH,"//*[@id='fof-layout']/div[1]/div/div[2]/div/div[2]/div/div[3]/div/div[2]/div[2]/div/div/div/div/div/table/tbody/tr[1]/td[2]").get_attribute("innerText")
            except selenium.common.exceptions.NoSuchElementException:
                web_date = "heihei"
            if web_date != value_date:
                NAVPU = df_update['单位净值'][fund_index]                                                                     # 获取基金单位净值
                CNAV = df_update['累计净值'][fund_index]                                                                      # 获取基金累计净值
                time.sleep(1)
                driver.find_element(By.XPATH, "//*[@id='fof-layout']/div[1]/div/div[2]/div/div[2]/div/div[3]/div/div[2]/div[1]/div[2]/button[2]").click()    # 点击上传净值
                time.sleep(1)
                driver.find_element(By.XPATH,"/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div[4]/form/div[1]/div[2]/div/span/span/div/input").click()  # 点击净值日期
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div/div/div/div/div[1]/div/input').send_keys(value_date)         # 填写净值日期
                except NoSuchElementException:
                    driver.find_element(By.XPATH,'/html/body/div[4]/div/div/div/div/div[1]/div/input').send_keys(value_date)         # 填写净值日期
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div[3]/button[1]').click()          # 随意点击
                except NoSuchElementException:
                    driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div[3]/button[1]').click()          # 随意点击
                time.sleep(1)
                driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div[4]/form/div[2]/div[2]/div/span/input').send_keys(NAVPU)                          # 填写单位净值
                driver.find_element(By.XPATH,'/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div[4]/form/div[3]/div[2]/div/span/input').send_keys(CNAV)                           # 填写累计净值
                time.sleep(1)
                driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/div[2]/div[3]/div/button[2]").click()    # 点击上传净值
                time.sleep(1)
                driver.find_element(By.XPATH, "/html/body/div[5]/div/div[2]/div/div[2]/div/div/div[2]/button[2]").click()    # 点击覆盖
                time.sleep(1)
            else:
                pass
            driver.close()
            time.sleep(1)
        else:
            print('\033[31m%s基金打开错误，好像是串行了，要改一下程序\033[0m' % fundname)
        time.sleep(1)
        driver.switch_to.window(driver.window_handles[0])                                                               # 切换到初始页面
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="fof-layout"]/header/div[1]/div[2]/div[2]/div/div/div/div/ul/li/div/span[1]').click()    # 随便点击
    except ValueError:
        if "已结束" not in fundname:
            if fundname in fund_list:
                pass
            else:
                print("\033[31m%s产品没有识别到净值，请检查《基金代码&名称对照表》，是否忘记填写净值，或产品已结束未标记\033[0m" % fundname)

# 登录火富牛
get_huofuniu_login(huofuniu_username, huofuniu_password)                                                                # 登录火富牛
df_update = pd.read_excel(update_excel_path)                                                                            # 读取基金净值表
time.sleep(5)
click_times = int(driver.find_element(By.XPATH,"//*[@id='fof-layout']/div[1]/div/div[2]/div/div[1]/div/div/div[4]/div[2]/ul/li[1]").get_attribute("innerText")[2:-2])/50  # 看一共有几页要点击几次翻页
if isinstance(click_times, int):
    click_times = int(click_times)-1
else:
    if click_times.is_integer():
        click_times = int(click_times)-1
    else:
        click_times = int(click_times)
# 上传净值
clicked=0
while clicked <= click_times:
    fund_count = int(driver.find_element(By.XPATH,"//*[@id='fof-layout']/div[1]/div/div[2]/div/div[1]/div/div/div[4]/div[1]/div/div/div/div/div/div/table/tbody").get_attribute("childElementCount"))+1
    for count in range(1, fund_count):
        update_value(count)
    driver.find_element(By.XPATH,'//*[@id="fof-layout"]/div[1]/div/div[2]/div/div[1]/div/div/div[4]/div[2]/ul/li[%s]'% (4+clicked)).click()  # 点击翻页
    time.sleep(2)
    driver.execute_script("window.scrollTo(0, 0);")  # 滚动到首屏
    clicked = clicked + 1
    time.sleep(2)
driver.close()
########################################################################################################################
# 下载产品净值到本地
print('开始下载净值')
session = requests.Session()                                                                                            # 保持request链接
# 备份已投产品净值
value_backuo_path = value_path + '_Backup'
time.sleep(1)
for file_name in os.listdir(value_path):
    source_file = os.path.join(value_path, file_name)
    destination_file = os.path.join(value_backuo_path, file_name)
    shutil.copy2(source_file, destination_file)
time.sleep(1)
# 删除现有记录
for filename in os.listdir(value_path):
    file_path = os.path.join(value_path, filename)
    # 如果是文件，直接删除
    if os.path.isfile(file_path):
        os.unlink(file_path)
time.sleep(1)
# 删除下载文件夹里所有文件
for filename in os.listdir(download_path):
    file_path = os.path.join(download_path, filename)
    # 如果是文件，直接删除
    if os.path.isfile(file_path):
        os.unlink(file_path)
time.sleep(1)
#  开始下载净值
device_id = '5b57f93282b7400e75f6c65c9e171cbe'
for _ in range(5):
    try:
        # 设置登录请求的参数
        login_url = "https://api.huofuniu.com/newgoapi/login"
        login_payload = {"username": huofuniu_username, "password": MD5.new(huofuniu_password.encode('utf-8')).hexdigest()}
        login_headers = {
            'authority': 'api.huofuniu.com',
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
            'cache-control': 'no-cache',
            'content-type': 'application/json;charset=UTF-8',
            'origin': 'https://mp.fof99.com',
            'pragma': 'no-cache',
            'referer': 'https://mp.fof99.com/',
            'sec-ch-ua': '"Not A(Brand";v="99", "Microsoft Edge";v="121", "Chromium";v="121"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'cross-site',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0',
            'x-device-id': device_id
        }
        # 发送登录请求
        response = session.post(login_url, headers=login_headers, json=login_payload, timeout=10)
        response.raise_for_status()
        data = response.json()
        # 提取token
        token = data['data']['token']
        # logger.info(f"{username} {data['msg']}")
        break  # 如果成功，退出重试循环
    except Exception as e:
        logger.error(f"登录异常: {e}")
def decrypt_des(encrypted_hex, key_str="ac68!3#1"):
    # 将密钥字符串转换为UTF-8编码的bytes
    key_bytes = key_str.encode('utf-8')
    # 初始化向量为密钥的UTF-8字节形式
    iv_bytes = key_bytes

    encrypted_bytes = unhexlify(encrypted_hex)

    # 创建DES Cipher对象，使用CBC模式
    cipher = DES.new(key_bytes, DES.MODE_CBC, iv_bytes)

    # 解密数据，并去除填充
    decrypted_bytes = unpad(cipher.decrypt(encrypted_bytes), DES.block_size)

    # 将解密后的bytes转换为UTF-8字符串
    decrypted_str = decrypted_bytes.decode('utf-8')

    return decrypted_str
def safe_request(method, url, session=None, **kwargs):
    try:
        if session is None:
            session = requests.session()
        # 设置代理
        # proxies = random_proxy()
        response = session.request(method, url, timeout=10, **kwargs)
        response.raise_for_status()
        # logger.info(f"{url}, ok")
        return response
    except Exception as e:
        logger.error(f"请求异常: {e}")
        raise e
def get_fund_value(session, token, device_id, fund_id, fund_name, value_path):
    # 获取数据
    url = f"https://pyapi.huofuniu.com/pyapi/fund/view?token={token}&fid={fund_id}&pt=1&shareToken="
    payload = {}
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Access-Token': token,
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Origin': 'https://mp.fof99.com',
        'Pragma': 'no-cache',
        'Referer': 'https://mp.fof99.com/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0',
        'X-DEVICE-ID': device_id
    }
    # 发送获取数据请求
    response = safe_request("GET", url, headers=headers, data=payload, session=session)
    data = response.json()['data']
    data['fund']['excess_prices'] = decrypt_des(data['fund']['excess_prices'])
    data['fund']['prices'] = decrypt_des(data['fund']['prices'])
    data['index']['prices'] = decrypt_des(data['index']['prices'])
    prices = json.loads(data['fund']['prices'])
    prices = sorted(prices, key=lambda x: x['pd'], reverse=True)
    df_nav = pd.DataFrame([[row['pd'], row['nav'], row['cnw'], row['cn'],row['pc'], row['drawdown']]for row in prices], columns=['净值日期', '单位净值', '累计净值', '复权净值', '收益率','回撤序列'])
    df_nav['净值日期'] = pd.to_datetime(df_nav['净值日期']).dt.date
    df_nav = df_nav.sort_values(by='净值日期', ascending=True)
    # df_nav.to_sql(fund_name, conn_fund_value, if_exists='replace', index=False)  # 管理人已发行产品存入sql
    df_nav.to_excel(value_path + "\%s.xlsx" % fund_name, index=False)  # 管理人已发行产品存入csv

for index in df_update.index:
    fund_id = df_update['火富牛基金代码'][index]
    fund_name = df_update['真实产品名称'][index]
    for _ in range(5):
        try:
            get_fund_value(session, token, device_id, fund_id, fund_name, value_path)
            break  # 如果成功，退出重试循环
        except:
            print('%s产品没有下载下来净值，index是%s' % (fund_name,index))
end_time = time.time()
elapsed_minutes = round((end_time - start_time) / 60, 2)
print('\033[32m%s只产品净值已更新，耗时%s分钟\033[0m' % (len(df_update), elapsed_minutes))
########################################################################################################################
# 净值写入投后分析
# 更新如期而至
sse = get_calendar("XSHG")                                                                                              # 获取中国股市交易日
df_fenxi = pd.read_excel(r"D:\SynologyDrive\投后分析表.xlsx",sheet_name="如期而至MOM",skiprows=1)[:40].dropna(how='all') # 读取投后分析表如期而至页面
app = xw.App(visible=False, add_book=False)                                                                             # 打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
exl = app.books.open(r"D:\SynologyDrive\投后分析表.xlsx")                                                                  # 打开要写入的文件
# 把净值填写到表格里
for index in df_fenxi.index:
    code = df_fenxi['账号'][index]                                                                                       # 获取产品代码
    name = df_fenxi['账户名'][index]                                                                                     # 获取产品名称
    if "已结束" not in name:
        fundsheet =pd.read_excel(r"D:\SynologyDrive\投后分析表.xlsx",sheet_name=code, skiprows=1).iloc[:, 1:]          # 读取产品详情页
        df_fund_value = pd.read_excel(r"D:\SynologyDrive\Data_Base\Invested_Fund_Value\%s.xlsx" % name)[['净值日期','单位净值','累计净值']]                 # 读取基金净值表
        date_new = df_fund_value['净值日期'].iloc[-1]                                                                     # 读取最新净值
        if date_new != fundsheet['日期'].iloc[-1]:
            netvalue = df_update.loc[df_update['产品简称'] == code]['单位净值'].values[0]                                      # 获取单位净值
            value = df_update.loc[df_update['产品简称'] == code]['累计净值'].values[0]                                         # 获取累计净值
            chengben = df_update.loc[df_update['产品简称'] == code]['持仓成本'].values[0]                                      # 获取持仓成本
            fene = df_update.loc[df_update['产品简称'] == code]['持仓份额'].values[0]                                          # 获取产品份额
            money = df_update.loc[df_update['产品简称'] == code]['当前动态权益'].values[0]                                      # 获取动态权益
            fenhong = df_update.loc[df_update['产品简称'] == code]['历史分红'].values[0]                                       # 获取历史分红
            earn = round(money - fundsheet['当前动态权益'].iloc[-1],2)                                                         # 较昨日涨跌幅
            churujin = fundsheet['出入金'].iloc[-1]
            fundsheet.loc[fundsheet.index.stop] = [date_new, netvalue, value, chengben, fene, money, fenhong, earn, churujin,'']         # 填写到产品详情表格内
            if round(money / fene, 4) != netvalue:
                print('%s净值好像有问题，检查一下，index是%s' % (name, index))
            fundsheet['日期'] = pd.to_datetime(fundsheet['日期']).dt.date
            exl.sheets[code].range("B3").value = fundsheet.values                                                           # 把DF写在excel里
exl.save()
exl.close()
# 获取某一频率收益率
def get_return(freq, date_new, name):
    if freq == 'week':
        start_date = date_new - datetime.timedelta(days=date_new.weekday())                                             # 获取本周第一天
    elif freq == 'month':
        start_date = date_new.replace(day=1)                                                                            # 获取本月第一天
    elif freq == 'quarter':
        start_date = date_new.replace(day=1, month=((date_new.month - 1) // 3) * 3 + 1)                                 # 获取本季度第一天
    elif freq == 'halfyear':
        start_date = date_new.replace(day=1, month=1 if date_new.month <= 6 else 7)                                     # 获取本半年第一天
    else:
        start_date = date_new.replace(day=1, month=1)                                                                   # 获取本年第一天
    trade_date = sse.valid_days(start_date=start_date, end_date=date_new).to_list()                                     # 获取本freq交易日
    first_day = trade_date[0].floor('D')                                                                                # 本freq交易日第一天去除频率
    first_day = first_day.tz_localize(None)                                                                             # 本freq交易日第一天去除时区
    if fundsheet['日期'][0] > first_day:
        first_day = fundsheet['日期'][0]
        freq_return = round(fundsheet['当前动态权益'].iloc[-1] - fundsheet['当前动态权益'][0],2)                             # 如果本freq第一个交易日还没投这个投顾，那么本freq第一个交易日就是投这个投顾的那天
    elif first_day == date_new:
        first_day = date_new
        freq_return = round(fundsheet['当前动态权益'].iloc[-1] - fundsheet['当前动态权益'].iloc[-2],2)                       # 昨日涨跌金额
    else:
        try:
            first_day = fundsheet['日期'][fundsheet[fundsheet['日期'] == first_day].index[0]-1]
        except KeyError:
            first_day = fundsheet['日期'][0]
        except IndexError:
            real_first_day = first_day
            first_day = trade_date[1].floor('D')                                                                        # 本freq交易日第一天去除频率
            first_day = first_day.tz_localize(None)                                                                     # 本freq交易日第一天去除时区
            print('%s的本%s初始日期没有净值，实际应为%s，用%s代替' % (name, freq, real_first_day, first_day))
        start_money = fundsheet['当前动态权益'][fundsheet['日期']==first_day].iloc[0]                                               # 上freq最后一个交易日动态权益                # 如果投资不足一freq，投资当天的动态权益
        freq_return = round(fundsheet['当前动态权益'].iloc[-1] - start_money,2)                                            # freq收益率
    # 判断有没有增资
    churujin = sum(fundsheet.set_index(fundsheet['日期'])['出入金'][first_day:date_new])                                                              # 截取本freq出入金
    freq_return = freq_return - churujin
    return freq_return
# 获取某一频率收益率
exl = app.books.open(r"D:\SynologyDrive\投后分析表.xlsx")                                                                  # 打开要写入的文件
for index in df_fenxi.index:
    code = df_fenxi['账号'][index]  # 获取产品代码
    name = df_fenxi['账户名'][index]  # 获取产品名称
    if "已结束" not in name:
        fundsheet = pd.read_excel(r"D:\SynologyDrive\投后分析表.xlsx", sheet_name=code, skiprows=1).iloc[:, 1:]        # 读取产品详情页
        date_new = pd.read_excel(r"D:\SynologyDrive\Data_Base\Invested_Fund_Value\%s.xlsx" % name)['净值日期'].iloc[-1]   # 读取最新净值
        week_return = get_return('week', date_new, name)                                                           # 本周收益率
        month_return = get_return('month', date_new, name)                                                         # 本月收益率
        quarter_return = get_return('quarter', date_new, name)                                                     # 本季收益率
        halfyear_return = get_return('halfyear', date_new, name)                                                   # 本半年收益率
        year_return = get_return('year', date_new, name)                                                           # 本年收益率
        index_temp = df_fenxi[df_fenxi['账号']==code].index[0]+3
        if len(fundsheet.columns[-1]) >5:
            name_temp = fundsheet.columns[-1][5:]
            try:
                fundsheet_temp = fundsheet.set_index("日期")
                fund_temp = pd.read_excel("D:\Database\Fund_Value\%s.xlsx" % name_temp, index_col="净值日期")
            except:
                print("【%s】产品净值不在数据库内" % name_temp)
        exl.sheets['如期而至MOM'].range("M%s" % index_temp).value = week_return                                           # 把DF写在excel里
        exl.sheets['如期而至MOM'].range("N%s" % index_temp).value = month_return                                          # 把DF写在excel里
        exl.sheets['如期而至MOM'].range("O%s" % index_temp).value = quarter_return                                        # 把DF写在excel里
        exl.sheets['如期而至MOM'].range("P%s" % index_temp).value = halfyear_return                                       # 把DF写在excel里
        exl.sheets['如期而至MOM'].range("Q%s" % index_temp).value = year_return                                           # 把DF写在excel里
exl.save()
exl.close()
app.kill()
