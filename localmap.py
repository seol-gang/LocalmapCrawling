from openpyxl import Workbook
import re
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

link_list = [
    #술집
    [
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=118&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=113&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=196&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=194&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=430&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=193&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=188&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=192&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=189&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=190&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=195&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=191&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=369&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=197&"
    ],
    #여가
    [
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=471&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=362&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=363&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=492&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=403&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=429&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=105&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=41&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=106&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=182&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=289&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=285&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=287&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=389&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=454&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=186&"
    ],
    #건강
    [
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=42&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=43&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=358&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=421&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=422&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=448&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=450&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=355&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=361&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=360&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=321&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=212&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=359&"
    ],
    #의료
    [
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=45&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=44&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=10&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=8&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=209&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=211&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=213&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=291&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=293&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=295&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=294&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=334&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=203&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=335&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=156&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=206&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=46&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=218&"
    ],
    #관광
    [
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=64&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=175&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=254&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=253&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=256&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=257&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=263&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=377&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=388&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=102&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=378&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=98&",
        "http://localmap.co.kr/web/splus/kmap/list.php?sigun=6310000&gugun=3700000&keyno=99&"
    ]
]

content_name = ['술집', '여가', '건강', '의료', '관광']

idx = 2
end_page = []
shop_link = []
parse_data = []
write_wb = Workbook()
write_ws = write_wb.active
write_ws.title = "남구"
write_ws['B1'] = '업종'
write_ws['C1'] = '회사이름'
write_ws['D1'] = '문의전화'
write_ws['E1'] = '지번주소'
write_ws['F1'] = '개업일자'
driver = webdriver.Chrome('chromedriver.exe')

def waitPage(xpath):
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, xpath)))
    except Exception:
        driver.refresh()
        return False
    return True

def getEndPageVal(link_list_):
    for link in link_list_:
        driver.get(link)
        try:
            while True:
                if waitPage('//*[@id="thema_wrapper"]/div[2]/div/div/div[1]/div[2]/font'): break
            a = driver.find_element_by_xpath('//*[@id="thema_wrapper"]/div[2]/div/div/div[1]/div[2]/font').text
        except Exception:
            continue
        end_page.append(int(re.findall('\d+', a)[1]))

def getContents_link(link_list_, name):
    global shop_link
    for idx, link in enumerate(link_list_):
        if idx < 8:
            continue
        for i in range(1, end_page[idx]+1):
            try:
                link_ = link + "&page=" + str(i)
                driver.get(link_)
                page = driver.page_source
                soup = BeautifulSoup(page, 'html.parser')
                shop_list = soup.find_all('h4')
                for idx, value in enumerate(shop_list):
                    try:
                        url = value.find('a').get('href')
                        if 'view.php' in url:
                            shop_link.append(value.find('a').get('href'))
                        else:
                            continue
                    except Exception:
                        continue
            except Exception:
                continue
            print(link_ + '\n')
            parseData(name)
            shop_link = []
            link_ = ''
        


def parseData(name):
    for link in shop_link:
        parse_dict = {}
        driver.get('http://localmap.co.kr' + link)
        while True:
            if waitPage('//*[@id="thema_wrapper"]/div[2]/div/div/div[2]/div[2]'): break
        parent = driver.find_element_by_xpath(
            '//*[@id="thema_wrapper"]/div[2]/div/div/div[2]/div[2]')
        table_box = parent.find_elements_by_tag_name('tr')
        for attr in table_box:
            t_attr = attr.text.split('\n')
            if '회사이름' in attr.text:
                parse_dict[t_attr[0]] = t_attr[1]
            elif '문의전화' in attr.text:
                parse_dict[t_attr[0]] = t_attr[1]
            elif '지번주소' in attr.text:
                parse_dict[t_attr[0]] = t_attr[1]
            elif '개업일자' in attr.text:
                parse_dict[t_attr[0]] = t_attr[1]
            else:
                continue
        parse_data.append(dict(parse_dict))
    saveDataToExcel(name)

def saveDataToExcel(name):
    global parse_data, idx
    for data in parse_data:
        write_ws['B'+str(idx)] = name
        write_ws['C'+str(idx)] = data['회사이름']
        try:
            write_ws['D'+str(idx)] = data['문의전화']
        except Exception:
            write_ws['D'+str(idx)] = ''
        write_ws['E'+str(idx)] = data['지번주소']
        write_ws['F'+str(idx)] = data['개업일자']
        idx += 1
    parse_data = []
    write_wb.save('./test.xlsx')
    

if __name__ == "__main__":
    #링크 페이지의 끝페이지 가져오기
    for link, name in zip(link_list, content_name):
        getEndPageVal(link)
        getContents_link(link, name)
        write_wb.save('./test.xlsx')
        end_page = []
    write_wb.save('./test.xlsx')