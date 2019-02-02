import requests
from lxml import html
import ocr
import Write
"""
                                            Created by ABHISHEK KOUSHIK B N 
                                                        AND
                                                       AKASH R
                                            on 01/02/2019
 """

usn_list = []
ia1 = []
ia2 = []
ia3 = []
ia4 = []
ia5 = []
ia6 = []
ia7 = []
ia8 = []
ea1 = []
ea2 = []
ea3 = []
ea4 = []
ea5 = []
ea6 = []
ea7 = []
ea8 = []
t1 = []
t2 = []
t3 = []
t4 = []
t5 = []
t6 = []
t7 = []
t8 = []
i = 1
while i <= 217:
    s = requests.Session()
    headers = {'Referer': 'http://results.vtu.ac.in/vitaviresultcbcs2018/index.php',
                   'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36',
                   'Upgrade-Insecure-Requests': '1',  'Cookie': 'PHPSESSID=u47uot7eg9j6eglqm951e3nfr7'
            , 'Connection': 'keep-alive'}
    image = s.get("http://results.vtu.ac.in/vitaviresultcbcs2018/captcha_new.php", headers=headers)
    with open("snap.png", 'wb') as file:
        file.write(image.content)
    cap = ocr.get_ocr("snap.png")
    USN = "1BI16CS"+str(format(i, '03d'))
    url = "http://results.vtu.ac.in/vitaviresultcbcs2018/resultpage.php"
    payload = {'lns': USN, 'captchacode': str(cap),
                   'token': 'YVY3aVVsT1F0dGNhalpxZFU1c0VuYjdsOER4VlRUay81alRYUUFucmRyekpHVWxaM2owby9PNGJYMlE2elREMUp6UkdOTk1IcXdQTnBVSkh4eFZiMGc9PTo67Sdlq3FpWDAYuCoX3rutjQ==',
                   'current_url': 'http://results.vtu.ac.in/vitaviresultcbcs2018/index.php'}
    page = s.post(url, data=payload, headers=headers)
    tree = html.fromstring(page.content)
    print("Sent USN:-1BI16CS"+str(format(i, '03d')))
    print("Sent Captcha:"+ocr.get_ocr("snap.png"))
    if "Invalid captcha code !!!" in page.text:
        print("Invalid captcha code !!!")
        continue
    else:
        i += 1
    if "Redirecting to VTU Results Site" in page.text:
        print("Alert:-Token Expired!!:Update new token in Payload")
        exit(2)
    if "University Seat Number is not available or Invalid..!" in page.text:
        ia1.insert(len(ia1), "-")
        ia2.insert(len(ia2), "-")
        ia3.insert(len(ia3), "-")
        ia4.insert(len(ia4), "-")
        ia5.insert(len(ia5), "-")
        ia6.insert(len(ia6), "-")
        ia7.insert(len(ia7), "-")
        ia8.insert(len(ia8), "-")
        ea1.insert(len(ea1), "-")
        ea2.insert(len(ea2), "-")
        ea3.insert(len(ea3), "-")
        ea4.insert(len(ea4), "-")
        ea5.insert(len(ea5), "-")
        ea6.insert(len(ea6), "-")
        ea7.insert(len(ea7), "-")
        ea8.insert(len(ea8), "-")
        t1.insert(len(t1), "-")
        t2.insert(len(t2), "-")
        t3.insert(len(t3), "-")
        t4.insert(len(t4), "-")
        t5.insert(len(t5), "-")
        t6.insert(len(t6), "-")
        t7.insert(len(t7), "-")
        t8.insert(len(t8), "-")
        continue
    if("Semester : 4" in page.text):
        imarks1 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div[3]')[0].text
        imarks2 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[3]/div[3]')[0].text
        imarks3 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[4]/div[3]')[0].text
        imarks4 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div[3]')[0].text
        imarks5 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[6]/div[3]')[0].text
        imarks6 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[7]/div[3]')[0].text
        imarks7 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[8]/div[3]')[0].text
        imarks8 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[9]/div[3]')[0].text
        emarks1 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/div[4]')[0].text
        emarks2 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[3]/div[4]')[0].text
        emarks3 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[4]/div[4]')[0].text
        emarks4 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div[4]')[0].text
        emarks5 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[6]/div[4]')[0].text
        emarks6 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[7]/div[4]')[0].text
        emarks7 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[8]/div[4]')[0].text
        emarks8 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[9]/div[4]')[0].text
        tmarks1 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/div[5]')[0].text
        tmarks2 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[3]/div[5]')[0].text
        tmarks3 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[4]/div[5]')[0].text
        tmarks4 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div[5]')[0].text
        tmarks5 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[6]/div[5]')[0].text
        tmarks6 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[7]/div[5]')[0].text
        tmarks7 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[8]/div[5]')[0].text
        tmarks8 = tree.xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[9]/div[5]')[0].text
        ia1.insert(len(ia1), imarks1)
        ia2.insert(len(ia2), imarks2)
        ia3.insert(len(ia3), imarks3)
        ia4.insert(len(ia4), imarks4)
        ia5.insert(len(ia5), imarks5)
        ia6.insert(len(ia6), imarks6)
        ia7.insert(len(ia7), imarks7)
        ia8.insert(len(ia8), imarks8)
        ea1.insert(len(ea1), emarks1)
        ea2.insert(len(ea2), emarks2)
        ea3.insert(len(ea3), emarks3)
        ea4.insert(len(ea4), emarks4)
        ea5.insert(len(ea5), emarks5)
        ea6.insert(len(ea6), emarks6)
        ea7.insert(len(ea7), emarks7)
        ea8.insert(len(ea8), emarks8)
        t1.insert(len(t1), tmarks1)
        t2.insert(len(t2), tmarks2)
        t3.insert(len(t3), tmarks3)
        t4.insert(len(t4), tmarks4)
        t5.insert(len(t5), tmarks5)
        t6.insert(len(t6), tmarks6)
        t7.insert(len(t7), tmarks7)
        t8.insert(len(t8), tmarks8)
    else:
        ia1.insert(len(ia1), "-")
        ia2.insert(len(ia2), "-")
        ia3.insert(len(ia3), "-")
        ia4.insert(len(ia4), "-")
        ia5.insert(len(ia5), "-")
        ia6.insert(len(ia6), "-")
        ia7.insert(len(ia7), "-")
        ia8.insert(len(ia8), "-")
        ea1.insert(len(ea1), "-")
        ea2.insert(len(ea2), "-")
        ea3.insert(len(ea3), "-")
        ea4.insert(len(ea4), "-")
        ea5.insert(len(ea5), "-")
        ea6.insert(len(ea6), "-")
        ea7.insert(len(ea7), "-")
        ea8.insert(len(ea8), "-")
        t1.insert(len(t1), "-")
        t2.insert(len(t2), "-")
        t3.insert(len(t3), "-")
        t4.insert(len(t4), "-")
        t5.insert(len(t5), "-")
        t6.insert(len(t6), "-")
        t7.insert(len(t7), "-")
        t8.insert(len(t8), "-")
print(len(ia1))
print(len(ia2))
print(len(ia3))
print(len(ia4))
print(len(ia5))
print(len(ia6))
print(len(ia7))
print(len(ia8))
print(len(ea1))
print(len(ea2))
print(len(ea3))
print(len(ea4))
print(len(ea5))
print(len(ea6))
print(len(ea7))
print(len(ea8))
print(len(t1))
print(len(t2))
print(len(t3))
print(len(t4))
print(len(t5))
print(len(t6))
print(len(t7))
print(len(t8))
for i in range(1, 218):
    usn_list.insert(len(usn_list), "1BI16CS"+str(format(i, '03d')))
Write.write_to_excel_internal(usn_list, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8)
Write.write_to_excel_external(usn_list, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8)
Write.write_to_excel_total(usn_list, t1, t2, t3, t4, t5, t6, t7, t8)
Write.close()
