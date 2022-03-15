import requests
from bs4 import BeautifulSoup
import re
import xlwt

book = xlwt.Workbook(encoding='utf-8')
sheet1 = book.add_sheet(u'Sheet1', cell_overwrite_ok=True)
sheet1.write(0, 0, '作品名')
sheet1.write(0, 1, '作者')
sheet1.write(0, 2, '尺寸')
sheet1.write(0, 3, '创作年代')
sheet1.write(0, 4, '成交价')
sheet1.write(0, 5, '拍卖时间')
sheet1.write(0, 6, '拍卖公司')
sheet1.write(0, 7, '拍卖会')
sheet1.write(0, 8, '说明信息')
row = 1

def write_file(name, author, size, create, price, time, company, meeting, info):
    l = len(name)
    print()
    for i in range(l):
        global row
        sheet1.write(row, 0, name[i])
        sheet1.write(row, 1, author[i])
        sheet1.write(row, 2, size[i])
        sheet1.write(row, 3, create[i])
        sheet1.write(row, 4, price[i])
        sheet1.write(row, 5, time[i])
        sheet1.write(row, 6, company[i])
        sheet1.write(row, 7, meeting[i])
        sheet1.write(row, 8, info[i])
        row += 1

def parseHTML(html):
    soup = BeautifulSoup(html,'html.parser')
    result = soup.find(name="ul", attrs={"class" :"imgList worksResult02 clearfix"})
    title = result.findAll("h3")
    price = []
    for e in result.findAll('em', text=re.compile(r'成交价')):
        price.append(e.parent)
    name = [t.a.text for t in title]
    price = [e.span.text for e in price]
    href = ['https://auction.artron.net' + t.a["href"] for t in title]
    # print(href)
    # print(name)
    # print(price)

    author = []
    size = []
    create = []
    time=[]
    company=[]
    meeting=[]
    info=[]
    for h in href:
        response = requests.get(h, headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.80 Safari/537.36'})
        res = response.text
        soup_in = BeautifulSoup(res, 'html.parser')
        detail = soup_in.find("table")
        detail = detail.find_all("td")

        if detail is not None and detail[0].a is not None:
            author.append(detail[0].a.text.strip())
        else:
            author.append('')
        if detail is not None and len(detail) > 1 and detail[1].em is not None:
            size.append(detail[1].em.text.strip())
        else:
            size.append('')
        if detail is not None and len(detail) > 3:
            create.append(detail[3].text.strip())
        else:
            create.append('')
        if detail is not None and len(detail) > 7 and detail[7].em is not None:
            time.append(detail[7].em.text.strip())
        else:
            time.append('')
        if detail is not None and len(detail) > 8 and detail[8].a is not None:
            company.append(detail[8].a.text.strip())
        else:
            company.append('')
        if detail is not None and len(detail) > 9 and detail[9].a is not None:
            meeting.append(detail[9].a.text.strip())
        else:
            meeting.append('')
        if detail is not None and len(detail) > 11:
            info.append(detail[11].text.strip())
        else:
            info.append('')

    write_file(name, author, size, create, price, time, company, meeting, info)





if __name__ == "__main__":

    headers = {
        'authority': 'auction.artron.net',
        'cache-control': 'max-age=0',
        'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="98", "Google Chrome";v="98"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.109 Safari/537.36',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'sec-fetch-site': 'same-site',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-user': '?1',
        'sec-fetch-dest': 'document',
        'referer': 'https://passport.artron.net/',
        'accept-language': 'zh-CN,zh;q=0.9',
        'cookie': 'Hm_lvt_851619594aa1d1fb8c108cde832cc127=1645981686; _dg_playback.63d6e55a16491bb2.29f3=1; _dg_abtestInfo.63d6e55a16491bb2.29f3=1; _dg_check.63d6e55a16491bb2.29f3=-1; _dg_antiBotFlag.63d6e55a16491bb2.29f3=1; _dg_antiBotInfo.63d6e55a16491bb2.29f3=10%7C%7C%7C3600; gr_user_id=62b760de-679f-454b-bef4-19891b64f519; _dg_attr.63d6e55a16491bb2.29f3=%7B%22userid%22%3A%223246327%22%7D; _dg_attr.63d6e55a16491bb2.365c=%7B%22userid%22%3A%223246327%22%7D; _guaid_status=1; gr_session_id_276fdc71b3c353173f111df9361be1bb=513ca5b8-5351-4042-b127-8e89efb787ae; gr_session_id_276fdc71b3c353173f111df9361be1bb_513ca5b8-5351-4042-b127-8e89efb787ae=true; _dg_id.63d6e55a16491bb2.365c=b298e8ec1b3d1219%7C%7C%7C1645981914%7C%7C%7C8%7C%7C%7C1645984980%7C%7C%7C1645984968%7C%7C%7C%7C%7C%7C676649e9668cb14d%7C%7C%7C%7C%7C%7C%7C%7C%7C1%7C%7C%7Cundefined; _at_pt_0_=3246327; _at_pt_1_=%E6%89%8B%E6%9C%BA%E7%94%A8%E6%88%B73246327; _at_pt_2_=e17899075bc2b8b53f83f2c4563e6f15; _dg_antiBotMap.63d6e55a16491bb2.29f3=202202280203%7C%7C%7C2; Hm_lpvt_851619594aa1d1fb8c108cde832cc127=1645985028; _dg_id.63d6e55a16491bb2.29f3=7ca6bf2477a30dc4%7C%7C%7C1645981686%7C%7C%7C14%7C%7C%7C1645985063%7C%7C%7C1645985028%7C%7C%7C%7C%7C%7C41415f75896d4564%7C%7C%7Chttps%3A%2F%2Fpassport.artron.net%2F%7C%7C%7C%7C%7C%7C1%7C%7C%7Cundefined',
    }

    params = (
        ('starttime', '2021-01-01'),
        ('endtime', '2021-12-31'),
    )

    for i in range(1, 11):
        response = requests.get(f'https://auction.artron.net/result/pmp-0-3-0-2-0-0-{i}/', headers=headers, params=params)
        # print(response.text)
        html = response.text
        parseHTML(html)
        print(i)

    book.save('result.xlsx')