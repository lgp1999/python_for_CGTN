import time
import requests
import re


def is_valid_date(date):
    '''
    判断是否是一个有效的日期字符串
    :param date: 匹配字符串
    :return: bool
    '''
    try:
        time.strptime(date, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def time_trans_stamp(date):
    '''
    将输入的日期转换为时间戳,用于页面查询
    :param date: 用户输入的日期
    :return:当前日期的凌晨0点的时间戳
    '''
    date1 = date + ' 00:00:00'
    # print(date1)
    time_array = time.strptime(date1, "%Y-%m-%d %H:%M:%S")
    time_stamp = int(time.mktime(time_array)) * 1000
    return time_stamp


def stamp_trans_time(date):
    # 将时间戳字符串转换成时间类型,返回两种时间格式
    time_stamp = int(int(date) / 1000)  # 接口取出来的时间戳不知道为啥多出三个零，去掉后时间才能正常转换
    loc_time = time.localtime(time_stamp)
    datetime1 = time.strftime("%Y-%m-%d %H:%M:%S", loc_time)
    datetime2 = time.strftime("%Y-%m", loc_time)
    datetime3 = time.strftime("%m-%d", loc_time)
    return [datetime1, datetime2, datetime3]


def get_name(text):
    '''
    利用正则表达式获取稿件内容中的记者
    :param text:
    :return:
    '''
    # rex1 = "（{1}(.+?)[记者]+ (.+?)）"
    rex1 = "[（|(]{1}[总台记者|总台央视记者|记者]+ [^（,(]*[)|）]"  # 正则匹配记者名字
    res1 = re.search(rex1, text)
    if res1:
        result1 = res1.group()
        return result1
    else:
        return None


def str_compare(str1, str2):
    '''
    对字符串进行匹配，str1是否在str2中存在
    :param str1:
    :param str2:
    :return:
    '''
    if str1 in str2:
        return True
    else:
        return False


def get_data(date, author_names):
    '''
    获取央视频app中时间链页面的某一天稿件数据
    :param date: 某一天的时间戳
    :param author_names: 要查找的所有记者名称
    :return: 稿件时间/阅读量/视频时长/标题/记者/稿件详情url
    '''
    url = 'http://api.cportal.cctv.com/api/rest/articleInfo/getScrollList'
    headers = {
        # headers，目前为空
    }
    page = 0
    count = 0  # 用来计算天数
    datas = {}
    for author_name in author_names:
        datas[author_name] = []
    # print(f"初始的：{datas}")
    while True:
        page += 1
        response = requests.get(url, params={'n': '20', 'version': '1', 'p': page, 'pubDate': date}, headers=headers)
        result = response.json()['itemList']
        # print(result)
        if result:
            for data in result:
                # 获取稿件详情，并取得记者名称
                item_id = data['itemID']  # 稿件id，用于获取稿件详情
                news_title = data['itemTitle']
                item_url = f'http://api.cportal.cctv.com/api/rest/articleInfo'
                news_response = requests.get(item_url, params={'id': item_id, 'cb': 'test.setMyArticalContent'})
                news_response.encoding = 'unicode_escape'  # 获取到的内容是Unicode编码，将其反编码显示汉字
                # print(news_response.encoding) # 调试 查看返回结果的编码格式
                news_result = news_response.text  # 获取接口返回的文本
                reporter_name = get_name(news_result)  # 获取文本匹配的（总台记者|总台央视记者）字符
                print(news_title , ' ' , reporter_name)  # 显示打印记者字段
                # 记者名字字符存在
                if reporter_name:
                    for author_name in author_names:
                        compare = str_compare(author_name, reporter_name)  # 查询记者名字是否在获取到的文本中
                        if compare:
                            news_url = data['detailUrl']
                            # news_title = data['itemTitle']
                            video_length = data['videoLength']  # 视频时长
                            put_date = data['pubDate']  # 发稿时间戳
                            put_time = stamp_trans_time(put_date)[0]  # 将发稿时间戳转换为北京时间
                            # print(news_title + '\t' + news_url + '\t' + put_time + '\t' + video_length)
                            # 获取稿件的阅读数量
                            view_url = 'http://nc.api.cportal.cctv.com/api/rest/clicknum'
                            view_params = {
                                'itype': 'news',
                                'id': item_id,
                                'cd': 'test.readCountDataSuccess'
                            }
                            view_response = requests.get(view_url, params=view_params)
                            view_result = view_response.json()['data']['vc']  # 阅读量数据
                            source = [put_time, view_result, video_length, news_title, reporter_name,
                                      news_url]  # 将一个稿件里的需要的数据放入一个列表
                            print(source)
                            datas[author_name].append(source)
                            # print(f"获取的：{datas}")
        if len(result) == 0:
            page = 0
            date += 86464000
            count += 1
        time.sleep(2)
        # 获取到数据为空时说明当天的稿件已经全部获取完毕，即该跳出循环
        if count == 3:
            break
        # 调试代码，只获取1页20条的数据
        # if page == 1:
        #     break
    # print(datas)
    return datas
