from urllib import request
import ast
import pandas as pd

host = 'https://restapi.amap.com/v3/weather/weatherInfo'
# key自己去https://lbs.amap.com/注册，配额内调用免费
key = 'key=' + 'XXXXXX'
# base:返回实况天气 all:返回预报天气
extensions = 'extensions=' + 'all'

# 城市编码表从https://lbs.amap.com/api/webservice/download中下载
city_list = ['310000', '320500', '321100', '370200', '321300', '610100', '500000']

contents = []
weather = pd.DataFrame(columns=('city', 'adcode', 'province', 'reporttime', 'date', 'week', 'dayweather',
                                'nightweather', 'daytemp', 'nighttemp', 'daywind', 'nightwind',
                                'daypower', 'nightpower'))
for city in city_list:
    adcode = 'city=' + city
    url = host + '?' + key + '&' + adcode + '&' + extensions
    response = request.urlopen(url)
    content = response.read().decode('utf-8')
    content_dict = ast.literal_eval(content)['forecasts'][0]
    df = pd.DataFrame(content_dict)
    df['casts'].apply(pd.Series)
    df = pd.concat([df.drop(['casts'], axis=1), df['casts'].apply(pd.Series)], axis=1)
    weather = weather.append(df, ignore_index=True)
weather.to_excel('weather.xlsx', index=False)
