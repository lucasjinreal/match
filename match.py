import os
import xlrd
from pprint import pprint
import sys
import datetime

def run(file):
    res_f_prefix = os.path.basename(file).split('.')[0]
    xls_file = xlrd.open_workbook(file)
    xls_sheet = xls_file.sheets()[0]

    rows = xls_sheet.nrows

    all_p = []

    all_man = []
    all_women = []

    result_f = open('result_{}.txt'.format(res_f_prefix), 'w', encoding='utf-8')

    # 6: 城市 3: 男女
    idx_city = 7
    idx_gender = 4
    idx_name = 2
    idx_wc_id = 3

    all_areas = []
    for i in range(rows):
        if i > 0:
            p = xls_sheet.row_values(i)
            if str(p[0]).isdigit():
                all_areas.append(p[idx_city])
                if p[idx_gender] == '男':
                    all_man.append(p)
                else:
                    all_women.append(p)
    print('男：{}人'.format(len(all_man)))
    result_f.write('男：{}人\n'.format(len(all_man)))
    print('\n\n')
    print('女：{}人'.format(len(all_women)))
    result_f.write('女：{}人\n\n'.format(len(all_women)))

    all_areas = list(set(all_areas))
    print('From: ', all_areas)
    result_f.write('From: {}'.format(', '.join(all_areas)))

    res = []
    # match algorithm

    un_res_women = []
    un_res_man = []

    matched_women = []
    matched_man = []
    # 1. find the same area man and woman first
    for women in all_women:
        # find the man from all_man
        if women not in matched_women:
            for m in all_man:
                if m[idx_city] == women[idx_city] and not m in matched_man:
                    one_pair = (women, m)
                    # remove m from all_man
                    matched_man.append(m)
                    matched_women.append(women)
                    res.append(one_pair)
                    break

    un_res_man = [i for i in all_man if i not in matched_man]
    un_res_women = [i for i in all_women if i not in matched_women]

    print('Matching result: ')
    for p in res:
        print('{}: {} --> 女:{} match 男:{}  all from: {}'.format(p[0][0], p[1][0],
                                                                p[0][idx_name], p[1][idx_name],
                                                                p[0][idx_city]))
        result_f.write('{}: {} --> 女:{} match 男:{}  all from: {}\n'.format(p[0][0], p[1][0],
                                                                           p[0][idx_name], p[1][idx_name],
                                                                           p[0][idx_city]))
        desc_0 = '\n'.join(p[0][4:])
        desc_1 = '\n'.join(p[1][4:])
        print(desc_1)
        result_f.write('\n匹配回复消息内容：\n'
                       '{}和{}, 欢迎你们成为有爱48小时CP。双方微信为{}, {}, 你们可以添加对方，主动一点哦！\n'
                       '来自{}的发言：{}\n'
                       '来自{}的发言：{}\n'
                       '希望你们在有爱群里继续完成打卡任务，任务完成后，将你们的故事分享给有爱后台，'
                       '我们会为你们制作一起颜值满满的推送哦！\n\n'.format(p[0][idx_name], p[1][idx_name],
                                                    p[0][idx_wc_id], p[1][idx_wc_id],
                                                    p[0][idx_name], desc_0,
                                                    p[1][idx_name], desc_1))
    print('====================== 格式二 ==========================')
    result_f.write('''入群请看群公告：

活动结果已出，群主公布CP名单，请各位小哥哥小姐姐认领爱的号码牌。

各自在自己的昵称前更改一下编号，好让对方认出你哦！然后你们可以互加好友，接收群主任务卡进行任务啦！\n''')
    i = 0
    for p in res:
        i += 1
        cp_id = '%03d' % i
        # get日期
        date = datetime.datetime.now().strftime('%M%d')
        result_f.write('【{}】 CP{} {} {} \n'.format(date, cp_id, p[0][idx_name], p[1][idx_name]))
    result_f.write('\n\n由于活动报名人数较多，以及地域关系，有些人没有匹配，大家可以关注【有爱白领圈】报名我们的推送活动，或者加入'
                   '相应的同城社群，我们目前开通了深圳，广州，北京，上海四地的本地群，添加小U微信拉入群聊')

    print('\n\n Unmatched: ')
    print('Man left: ')
    result_f.write('\n\n未匹配男： \n{}'.format('\n'.join([', '.join(i) for i in un_res_man])))
    pprint(un_res_man[:5])
    print('\n\nWomen left:')
    pprint(un_res_women[:5])
    result_f.write('\n\n未匹配女：\n {}'.format('\n'.join([', '.join(i) for i in un_res_women])))
    print('Done')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Pls provide xls.')
    else:
        run(sys.argv[1])
