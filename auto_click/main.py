import http.cookiejar as cookielib

from selenium import webdriver

driver = webdriver.Chrome()
driver.get("https://www.imooc.com/")
cj = cookielib.MozillaCookieJar('cookies.txt')
cj.load()
for c in cj:
    c_dict = {
        "domain": c.domain,
        "hostOnly": False,
        "httpOnly": False,
        "name": c.name,
        "path": c.path,
        "secure": c.secure,
        "session": True,
        "storeId": "0",
        "value": c.value,
    }
    driver.add_cookie(c_dict)
driver.refresh()
