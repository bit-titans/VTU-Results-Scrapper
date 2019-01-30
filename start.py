import ocr
from selenium import webdriver
from PIL import Image
import Write
driver = webdriver.Chrome("C:\\Users\\akash\\Downloads\\chromedriver_win32\\chromedriver.exe")
i = 1
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
while i <= 217:
    try:
        driver.get("http://results.vtu.ac.in/vitaviresultcbcs2018/index.php");
        usn = driver.find_element_by_name('lns')
        usn.send_keys("1BI16CS"+str(format(i, '03d')))
        print("Sent USN:-1BI16CS"+str(format(i, '03d')))
        driver.save_screenshot("snap.png")
        image = Image.open("snap.png")
        img = driver.find_element_by_xpath('//*[@id="raj"]/div[1]/div[3]/img')
        image = image.crop((int(img.location['x']), int(img.location['y']), int(img.location['x'])+int(img.size['width']), int(img.location['y'])+int(img.size['height'])))  # defines crop points
        image.save("snap.png", 'png')
        captcha = driver.find_element_by_name("captchacode")
        print("Sent Captcha:"+ocr.get_ocr("snap.png"))
        captcha.send_keys(ocr.get_ocr("snap.png"))
        submit = driver.find_element_by_id("submit")
        submit.submit()
        try:
            if(driver.switch_to.alert.text == "Invalid captcha code !!!"):
                pass
        except:
            i += 1
        sem = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[1]/b')
        if(sem.text == "Semester : 4"):
            imarks1 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div[3]')
            imarks2 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[3]/div[3]')
            imarks3 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[4]/div[3]')
            imarks4 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div[3]')
            imarks5 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[6]/div[3]')
            imarks6 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[7]/div[3]')
            imarks7 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[8]/div[3]')
            imarks8 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[9]/div[3]')
            emarks1 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/div[4]')
            emarks2 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[3]/div[4]')
            emarks3 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[4]/div[4]')
            emarks4 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div[4]')
            emarks5 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[6]/div[4]')
            emarks6 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[7]/div[4]')
            emarks7 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[8]/div[4]')
            emarks8 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[9]/div[4]')
            tmarks1 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[2]/div[5]')
            tmarks2 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[3]/div[5]')
            tmarks3 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[4]/div[5]')
            tmarks4 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[5]/div[5]')
            tmarks5 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[6]/div[5]')
            tmarks6 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[7]/div[5]')
            tmarks7 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[8]/div[5]')
            tmarks8 = driver.find_element_by_xpath('//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/div/div[9]/div[5]')
            ia1.insert(len(ia1), imarks1.text)
            ia2.insert(len(ia2), imarks2.text)
            ia3.insert(len(ia3), imarks3.text)
            ia4.insert(len(ia4), imarks4.text)
            ia5.insert(len(ia5), imarks5.text)
            ia6.insert(len(ia6), imarks6.text)
            ia7.insert(len(ia7), imarks7.text)
            ia8.insert(len(ia8), imarks8.text)
            ea1.insert(len(ea1), emarks1.text)
            ea2.insert(len(ea2), emarks2.text)
            ea3.insert(len(ea3), emarks3.text)
            ea4.insert(len(ea4), emarks4.text)
            ea5.insert(len(ea5), emarks5.text)
            ea6.insert(len(ea6), emarks6.text)
            ea7.insert(len(ea7), emarks7.text)
            ea8.insert(len(ea8), emarks8.text)
            t1.insert(len(t1), tmarks1.text)
            t2.insert(len(t2), tmarks2.text)
            t3.insert(len(t3), tmarks3.text)
            t4.insert(len(t4), tmarks4.text)
            t5.insert(len(t5), tmarks5.text)
            t6.insert(len(t6), tmarks6.text)
            t7.insert(len(t7), tmarks7.text)
            t8.insert(len(t8), tmarks8.text)
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
    except:
        print(driver.switch_to.alert.text)
        if (driver.switch_to.alert.text == "Invalid captcha code !!!"):
            driver.switch_to.alert.accept()
        try:
            if driver.switch_to.alert.text == "University Seat Number is not available or Invalid..!":
                driver.switch_to.alert.accept()
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
                i += 1

        except:
            pass
for i in range(1, 218):
    usn_list.insert(len(usn_list), "1BI16CS"+str(format(i, '03d')))
Write.write_to_excel_internal(usn_list, ia1, ia2, ia3, ia4, ia5, ia6, ia7, ia8)
Write.write_to_excel_external(usn_list, ea1, ea2, ea3, ea4, ea5, ea6, ea7, ea8)
Write.write_to_excel_total(usn_list, t1, t2, t3, t4, t5, t6, t7, t8)
Write.close()
