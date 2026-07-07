"""
Author:hippieZhou
Date:20190608
Description: Get BaiDuYun shared link's Code
"""
import datetime
from time import sleep

from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By

from util.util_web_driver import *
from util.util_file import *

VERSION = "VERSION 1.0.0"

cur_css_id = 0
dict_page: list = []
dict_css: dict = {}
dict_image: dict = {}

root_dir = "D:/0000_学习资料/books/"  # "./books/"
root_dir = "K:/usb1_1/0000_学习资料/books/"
root_dir = "D:/wksp/"
root_dir = "/mnt/readme/0000_书籍处理/0000_books/"
root_dir = "M:/0000_书籍处理/0000_books/"
def save_pdf_page(path: str, base_url: str, cur_page: int, fileMark, fileName):
    headers = {'Referer': 'https://ssj.sslibrary.com/',
               'Host':'pdfssj.sslibrary.com',
               'referer':'https://ssj.sslibrary.com',
               'Sec-Fetch-Dest': 'empty',
               'Sec-Fetch-Mode': 'cors',
               'Sec-Fetch-Site': 'same-site',
               'User-Agent':'Mozilla/5.0 (X11; Linux x86_64; rv:126.0) Gecko/20100101 Firefox/126.0'}

    the_url = base_url + "&cpage=" + str(cur_page)
    response = requests.get(the_url, headers=headers)
    if response.status_code == 200:
        file_name = path + "/" + fileMark + "_" + "{:0>4d}".format(cur_page) + ".pdf"
        f = open(file_name, 'wb')
        f.write(response.content)
        f.close()


if __name__ == "__main__":
    book_driver = BookDriver()

    book_driver.get_undetected_firefox_driver(1, "D:/Tools/geckodriver-v0.36.0-win32/geckodriver.exe", "C:/Users/acewa/.cache/selenium/firefox/win64/138.0.3/firefox.exe") # "D:/Tools/Mozilla Firefox/firefox.exe")

    book_driver.get("https://www.zhipin.com")

    # 再次注入以防万一 (某些网站会在 iframe 或新标签页中检测)
    book_driver.execute_script("""
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        });
    """)

    """
    ## 原来用这种账号密码的登录方式，密码过期了
    book_driver.get_driver().find_element("name", "username").send_keys("gdfgscs")
    book_driver.get_driver().find_element("name", "password").send_keys("gdfgs@2022")
    sleep(30)    
    # 改用下面的cookie登录方式
    """

    book_driver.add_zhaopin_cookie()
    # 改用上面的cookie登录方式, 不用密码了

    url = "https://www.zhipin.com"

    book_driver.get("https://www.zhipin.com")

    ddd = book_driver.get_undetected_firefox_driver().find_element(By.XPATH, "/html/body/div[1]/div[1]/div[1]/div[3]/ul/li[2]/a")

    ddd.click()

    sleep(30)



    books_result = "\n\n"
    books_result += datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    books_result += "\n\n"

    file_object = open("图书馆在借图书.txt", 'a', encoding='utf-8')  # 创建一个文件对象，也是一个可迭代对象
    try:
        file_object.writelines(books_result)  # 结果为str类型
    finally:
        file_object.close()

    for theId in range(0, 8):

        # 清除所有 Cookie
        book_driver.get_undetected_firefox_driver().delete_all_cookies()

        print("进入登录页面 authorize")

        url = "https://www.zhipin.com"
        book_driver.get(url)
        sleep(2)

        

        login_ok = False

        while not login_ok:
            ddd = book_driver.get_driver().find_element( By.ID, "username")
            ddd.clear()
            ddd.send_keys(id_list[theId])
            ddd = book_driver.get_driver().find_element(By.ID, "password")
            ddd.clear()
            ddd.send_keys(pass_list[theId])

            # 通过 class 定位
            try:
                ddd = book_driver.get_driver().find_element(By.ID, "agreement-username-login")
                if ddd.is_selected():
                    print("agreement-复选框已被选中")
                else:
                    print("agreement-复选框未被选中-点击")
                    ddd.click()
            except NoSuchElementException:
                print("没有找到 agreement-复选框")

            # ddd = book_driver.get_driver().find_element(By.ID, "agreement-username-login")
            # if ddd.is_selected():
            #     print("agreement-复选框已被选中")
            # else:
            #     print("agreement-复选框未被选中-点击")
            #     ddd.click()

            ddd = book_driver.get_driver().find_element(By.ID, "inputCaptcha")
            ddd.send_keys("")

            wait_time = 15
            while True:
                print(f"验证码输入时间   = {wait_time}")
                sleep(1)

                ddd = book_driver.get_driver().find_element(By.ID, "inputCaptcha")
                ddd_value = ddd.get_attribute("value")
                if len(ddd_value) == 4:
                    if wait_time > 3:
                        wait_time = 3
                        print(f"验证码输入完成2={wait_time}")
                    else:
                        print(f"验证码输入完成1={wait_time}")
                else:
                    wait_time = 15
                    print(f"验证码输入未完成={wait_time}")

                wait_time -= 1

                if wait_time <= 0:
                    print(f"验证码超时={wait_time}")

                    try:
                        ddd = book_driver.get_driver().find_element(By.ID, "login")
                        # 将焦点移动到控件
                        # 使用 JavaScript 将焦点移动到控件
                        print("将焦点移动到控件 login")
                        book_driver.get_driver().execute_script("arguments[0].focus();", ddd)
                        # ddd.click()
                        sleep(5)
                    except NoSuchElementException:
                        pass
                    else:
                        print("点击登录")
                        ## book_driver.get_driver().execute_script("arguments[0].click();", ddd)
                        # ddd = book_driver.get_driver().find_element(By.ID, "oauthLoginForm")
                        # 提交表单
                        # ddd.submit()

                        # print("ddd.click()")
                        # ddd.click()

                    ddd = book_driver.get_driver().find_element(By.ID, "inputCaptcha")
                    ddd_value = ddd.get_attribute("value")
                    if len(ddd_value) == 4:
                        break
                    else:
                        ddd.send_keys("")


            ns = 'document.querySelector("#activeTab").value = "home"; aaa = document.querySelector("#login"); \n aaa.click();'
            print("执行登录:  " + ns)
            res = book_driver.execute_script(ns)


            # 如果没有 login 按钮，则跳过
            try:
                print("查询是否还有 login 控件")
                sleep(5)
                ddd = book_driver.get_driver().find_element(By.ID, "login")
                # 将焦点移动到控件
                # 使用 JavaScript 将焦点移动到控件
                book_driver.get_driver().execute_script("arguments[0].focus();", ddd)
                # ddd.click()
            except NoSuchElementException:
                login_ok = True
                print("没有 login 控件，登录成功")
            else:
                login_ok = False
                print("还有 login 控件，登录失败")

        try:
            print("进入借书页面 borrowBooks")
            book_driver.get("https://www.library.sh.cn/myLibrary/borrowBooks")
            sleep(3)
        except:
            pass

        books_result = "\n\n"

        books = book_driver.get_driver().find_elements(By.CLASS_NAME, "infoItem")

        for book in books:
            books_result += id_list[theId] + "\t"

            bookCnt = book.find_element(By.CLASS_NAME, "content")

            bookInfos = bookCnt.find_elements(By.CLASS_NAME, "bookInfo")

            text0 = bookInfos[0].text
            books_result += text0.split(":")[1].split(",")[1].strip()
            books_result += "\t"

            text0 = bookInfos[1].text
            books_result += text0.split(":")[1].strip()
            books_result += "\t"

            text0 = bookInfos[2].text
            books_result += text0.split(":")[1].strip()
            books_result += "\t"

            text0 = bookInfos[3].text
            books_result += text0.split(":")[1].strip()
            books_result += "\t"

            name = bookCnt.find_element(By.CLASS_NAME, "c-title").find_element(By.CLASS_NAME, "name").text
            books_result += name + "\n"

        print("保存图书列表")
        file_object = open("图书馆在借图书.txt", 'a', encoding='utf-8', newline = "\n")  # 创建一个文件对象，也是一个可迭代对象
        try:
            file_object.writelines(books_result)  # 结果为str类型
        finally:
            file_object.close()

        try:
            print("进入我的页面 myLibrary")
            book_driver.get("https://www.library.sh.cn/service/myLibrary")
        except:
            pass
        sleep(4)

        print("点击退出 u-book-bn-logout")
        ddd = book_driver.get_driver().find_element(By.CLASS_NAME, "u-book-bn-logout")
        book_driver.get_driver().execute_script("arguments[0].click();", ddd)
        sleep(1)

        print("确认退出 u-book-bn-logout")
        ddd = book_driver.get_driver().find_element(By.CLASS_NAME, "v-card__actions")
        ttt = ddd.find_elements(By.CLASS_NAME, "v-btn")
        book_driver.get_driver().execute_script("arguments[0].click();", ttt[1])
        sleep(4)

        # 清除所有 Cookie
        book_driver.get_driver().delete_all_cookies()
        sleep(4)

    print("全部完成，等待300秒后退出")
    sleep(300)
    print("退出")
    book_driver.quit()
    
