# json to excel
# coding:utf8
import json
import openpyxl
import re

# -------------------- 標題清洗pattern -----------------------
# 取代為""   【xxx】：
pattern_title = re.compile(r'【圖輯】');

# -------------------- 內文清洗pattern -----------------------
# 取代為""   xxx編輯xxx
pattern_intern = re.compile(r'((\(|（).{0,6}編輯|(\(|（)譯者).{2,8}(）|\))');
# 取代為""   特定詞語+連結
pattern_fb = re.compile(r'(臉書全文：|直播畫面：|聲明全文：|原文連結：|連結：|請點我☛|資料來源：|同場加映：|影片：|推特全文：|觀看影片請點|如無法觀看影片|更多精彩照片看這裡|請點此連結|APP用戶如欲觀看|App用戶看不到影片|APP用戶觀看影片|手機APP用戶請點此觀看​|用戶請點：|Dcard原文連結：|手機用戶請點|報名網址|PTT原文：|手機用戶無法觀看影片請點|影片來源|影片連結|影片請點此|點我看更多|更多資訊請洽官方粉絲團).*(http://|https://|www)[a-zA-Z0-9\\./_)）]+');
# 取代為""   所有連結
pattern_conn = re.compile(r'(http://|https://|www)[a-zA-Z0-9\\./_]+')
# 取代為""   特定語句
pattern_sentence = re.compile(
    r'(NUTS圖片來源／|【點我看更多詳細內容】|喝酒不開車、開車不喝酒！|未成年請勿飲酒，飲酒過量，有礙健康！|莫逞一時樂，遺害百年身！|拒絕毒品　珍惜生命|健康無價　不容毒噬)|圖／|影／|文/|文字／|編輯／|美術設計／|攝影／|模特／|文章來源／|(請鎖定《三立新聞網》|健康有方)')
# 取代為""   xxxx/xx報導
pattern_report = re.compile(r'[\u4e00-\u9fa5]{2,4}(／|/)[\u4e00-\u9fa5]{2,4}(報導)')
# 取代為""   ✿ (多為裝飾符)
pattern_flower = re.compile(r'(✿|❤|\(\)|（）|　|◆|＊)');

# -------------------- pattern宣告結束 -----------------------



workbook = openpyxl.Workbook()
sheet = workbook.worksheets[0]
listtitle = ["title", "date", "author", "text", "url", "tags", "type_list", "source", "views", "share", "like"]
sheet.append(listtitle)
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
ILLEGAL_CHARACTERS_EMOJI = re.compile(u'[\U00010000-\U0010ffff]')
with open("D:/news/setn10000.json", "r", encoding='utf8') as jsonfile:
    data = json.load(jsonfile)
    for i in data.keys():
        tags = ""
        if data[i]['tags']:
            for tagstring in data[i]['tags']:
                tags = tags + tagstring + "、"
            tags = tags[0:-1]
            # print(tags)

        type_list = ""
        if data[i]['type_list']:
            for type_liststring in data[i]['type_list']:
                type_list = type_list + type_liststring + "、"
            type_list = type_list[0:-1]
            # print(type_list)

# -------------------- 標題清洗 -----------------------
        title = data[i]['title'];
        if re.search(pattern_title, title):
            title = pattern_title.sub('', title);

# -------------------- 內文清洗 -----------------------
        content = data[i]['text'];
        # -------------------- 清除後續 -----------------------
        # 清除 "※"後的所有內容
        check_sup = content.__contains__('※');
        if check_sup:
            split_sup = content.split('※');
            content = split_sup[0];
        # 清除 "░"後的所有內容
        check_mark1 = content.__contains__('░');
        if check_mark1:
            split_mark1 = content.split('░');
            content = split_mark1[0];
        # 清除 "【更多"後的所有內容
        check_mark3 = content.__contains__('【更多');
        if check_mark3:
            split_mark3 = content.split('【更多');
            content = split_mark3[0];
        # 清除 "延伸閱讀"後的所有內容
        check_extend = content.__contains__('延伸閱讀');
        if check_extend:
            split_extend = content.split('延伸閱讀');
            content = split_extend[0];
        # 清除 "更多完整內容直播"後的所有內容
        check_moredetail = content.__contains__('更多完整內容直播');
        if check_moredetail:
            split_moredetail = content.split('更多完整內容直播');
            content = split_moredetail[0];
        # 清除 "作者："後的所有內容
        check_authorInText = content.__contains__('作者：');
        if check_authorInText:
            split_authorInText = content.split('作者：');
            content = split_authorInText[0];

        # -------------------- 取代單個 -----------------------
        check_intern = re.search(pattern_intern, content);
        if check_intern:
            content = pattern_intern.sub('', content);

        check_fb = re.search(pattern_fb, content);
        if check_fb:
            content = pattern_fb.sub('', content);

        check_conn = re.search(pattern_conn, content);
        if check_conn:
            content = pattern_conn.sub('', content);

        check_sentence = re.search(pattern_sentence, content);
        if check_sentence:
            content = pattern_sentence.sub('', content);

        check_report = re.search(pattern_report, content);
        if check_report:
            content = pattern_report.sub('', content);

        check_flower = re.search(pattern_flower, content);
        if check_flower:
            content = pattern_flower.sub('', content);

# -------------------- 清洗結束 -----------------------




        row = [title, data[i]['time'].split('_')[0], data[i]['author'], content,
               data[i]['url'] + " ", tags, type_list, data[i]['source'], data[i]['like'], data[i]['share'],
               data[i]['views']];
        count = 0
        for i in row:
            row[count] = ILLEGAL_CHARACTERS_RE.sub(r'', row[count])
            row[count] = ILLEGAL_CHARACTERS_EMOJI.sub(r'', row[count])
            count += 1
        sheet.append(row)
workbook.save("D:/news/setn10000.xlsx")
