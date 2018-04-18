# -*- coding:utf-8 -*-
"""
@Author: lamborghini
@Date: 2018-04-08 19:21:01
@Desc: 
https://www.aliexpress.com/category/100005652/bakeware.html?
https://www.aliexpress.com/category/100005652/bakeware/1.html
https://www.aliexpress.com/category/1541/home-storage-organization.html?
https://www.aliexpress.com/category/1541/home-storage-organization/2.html
"""

import re
import xlwt

from pubcode import misc, pubcrawler
from bs4 import BeautifulSoup


class CAliExpress(pubcrawler.CPubCrawler):
    m_Flag = misc.Time2Str(timeformat="%Y-%m-%d")
    # m_Flag = "aliexpress"
    m_Url = "https://www.aliexpress.com/category/100005652/bakeware/"
    m_DelList = ["(", ")", " ", "Orders", "Order"]
    m_MyHeader = {
        "Accept-Encoding"   :"gzip, deflate, br",
        "Cookie"            :"ali_apache_id=10.183.74.35.1523185826444.214859.2; xman_us_f=x_l=1&x_locale=en_US&no_popup_today=n&x_user=CN|Xiao|Hao|ifm|1613424031&zero_order=y&last_popup_time=1523190444694; aep_usuc_f=region=US&site=glo&b_locale=en_US&isb=y&x_alimid=1613424031&c_tp=USD; intl_common_forever=NLZt7i8/lWAcxQbMVcuT49ivdubYsbTKUck3sNzxaDgwMK/S/JpSjw==; xman_f=nwlNGqfGCvuFlev4bj90sKUV1l1S/kdo5g2MZwT90w6dtu0yUMFGwxoyWBOVi8xfDFrir5yw665EOGp6UNAoAiqM9mZ/TaJ2UTrB3SCUp+7CC9vMqjqq1GrmWiF0Is701YOLvAdsAC1dBbzb+jvkW1slt+CigzDJEeeaxljJB3BTAFyKrE2XBCkZO6P6nQdRh7lFHqsQlwcuCNxHLUW7DhNmFKveAkpueJOkDQK0GS1ZLFKwp7UrmDKm+Siw57SekC7vxpBjv3IHZXh413y1i5v0vzOT3h3/pJd79L0SyWyVxqYw0grS8j61mAxQdaB5S98cvJAlbRs8qhSNXzvBDEKK+yDbQnDeewmS1CWWPnkqHgoJfaR35Rsl69Xm/GdXp6VsLPAfAk5ZHJyT/WHc9w==; _ga=GA1.2.439760157.1523185830; cna=WLpUEyr2mAUCAQ6Taj0Ha7fn; isg=BPr6EOKqdLj2nfjTPrrcgA1nSCMWi2u9sP_CigTzvw1U95ox7DvOlcCFQwMr_PYd; _uab_collina=152318659767449354775796; aep_history=keywords%5E%0Akeywords%09%0A%0Aproduct_selloffer%5E%0Aproduct_selloffer%0932766389425%0932848000412%0932825118129%0932790447927%0932856751924%0932740190143%0932840343211; _umdata=C234BF9D3AFA6FE7F178821B34DFC1AC406FCFCEC38E1772C391EF5D50DC8ABA880FF6309BF5ADE7CD43AD3E795C914C769B7B0FCAFB12BCB76AA2BEED071262; ali_beacon_id=10.183.74.35.1523185826444.214859.2; ali_apache_track=mt=1|ms=|mid=cn11731031aiuae; aep_common_f=9L8ombdfGfz5oe7HF12G2srs+VcZ2BzU75T7uU4hdjFgsEpSPomyuw==; __utma=3375712.439760157.1523185830.1523190446.1523190446.1; __utmz=3375712.1523190446.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); l=AgsLWpCTAYM/AhNWXZsNt0R-m6X1yB8i; JSESSIONID=FA8LSJ1W-C88UDXF8T4RQ33N57VVX3-TIL5VUFJ-P4E9; acs_usuc_t=acs_rt=ade6fc239bc345e783d768fda5775e75&x_csrf=lry_jv57de4l; intl_locale=en_US; xman_t=gUACFEr4O6sZbEXL5CTKYUOzIiWzrIq7BQd8m0feIHVZgz6LjvISVQok0m6dyYPHD8LCvgHijfz0Go8VRZ1l2VmkQY0e6+PM5rBdW0PYmQFD03fPX2/v8D5mDsQrKuGHdfXC1E9IoWnGj/lh3b5c9lFTTTe3OaCDAubuhmNMe10x86mJ+5Q0LGDZ4qTu0k730gzBScvO4LLzGZB6pNnpdKB2wjief6KNczt8KOjWFf2fGsvPCKP8HBpANkVlR19QKpgv8ZKaOJvP+eMeEiv0AdNgF/1aTNmv4h5oww4ARss3by7A9/Sf0CqNJOzl28TSOyY0UIeCL4gHAN7VfP7kbbk4gxBKFmAOEHUtyhoxSa5RaAqikc12scqxBT9fCXZdFX7ZcwfcEfOkZPrN3d6+3uj2ijxPvxg7h9TnVVXaLMB7M14M4QLCl4P52gF0K7aU1YSouUKFlKRkosq//KCAjRBfgWAo55uZBvyIk5amcSH0/kVBN502/Hy5fTqEO8UQrXCPZw5Tkc2W4aCDSshWKxLiCeTRGanyU+u/GWvgazAFLqGNrDz0VsxeHkr9RogSxJ5j1ATA40oUbk2cmkNBIY/OurD9SruDyxkNCMpVrWEJxOMtIQ4m1w==; _mle_tmp0=eNrz4A12DQ729PeL9%2FV3cfUxiKrOTLFScnO08An2MgzXdbawCHWJcLMIMQkKNDb2MzUPC4sw1g3x9DENC3Xz0g0wcbVU0kkusTI0NTI2MTY3NDM3MTTWSUxGE8itsDKojQIAN6gcRw%3D%3D; _gid=GA1.2.1030077585.1523437169; ali_apache_tracktmp=W_signed=Y; _hvn_login=13; xman_us_t=x_lid=cn11731031aiuae&sign=y&x_user=0ilRDWcCUJ57KKdf6Yxqt5kqhAhvCzcKCLPk90CuGuQ=&ctoken=o6w625aazgv7&need_popup=y&l_source=aliexpress; aep_usuc_t=ber_l=A0",
        "Host"              :"www.aliexpress.com",
        "Referer"           :"https://www.aliexpress.com/category/100005652/bakeware/2.html?site=glo&g=y&tag=&smToken=f0331d0f98b44bc299ee98b754504a79&smSign=tvTfy3yQR1mgpYZaQstmkg%3D%3D",
        "Upgrade-Insecure-Requests" :"1",
    }
    m_DebugPrint = True

    def _CustomInit(self):
        for x in range(1, 101):
            pageurl = self.m_Url + str(x) + ".html"
            if pageurl in self.m_DoneInfo:
                continue
            dPageInfo = {
                "priority"  :   0,
                "time"      :   misc.GetSecond(),
            }
            self.m_WaitingUrl[pageurl] = dPageInfo


    def SetUrlStart(self, url):
        self.m_Url = url
        self.Start()
        self.MySorted()


    def MySorted(self):
        lst = []
        for url, info in self.m_DoneInfo.items():
            if not "xxoo" in info:
                continue
            xxoo = info["xxoo"]
            lst.append((url, xxoo))
        aa = sorted(lst, key=lambda x:x[1], reverse=True)
        lstTitle = ["物品名", "价格", "反馈数", "订单数", "差值", "链接"]
        lstResult = [lstTitle]
        for url, xxoo in aa:
            dInfo = self.m_DoneInfo[url]
            sName = dInfo["name"]
            sPrice = dInfo["price"]
            iFeedBack = dInfo["feedback"]
            iOrder = dInfo["order"]
            tInfo = [sName, sPrice, iFeedBack, iOrder, xxoo, url]
            lstResult.append(tInfo)
        self.Save2Execl(lstResult)


    def Save2Execl(self, lstAnswer):
        sFileName = "%s.xls" % misc.Time2Str(timeformat="%Y%m%d%H%M%S")
        oBook = xlwt.Workbook()

        alignment_center = xlwt.Alignment()
        alignment_center.horz = xlwt.Alignment.HORZ_CENTER
        alignment_center.vert = xlwt.Alignment.VERT_CENTER

        oTitleStyle = xlwt.XFStyle()
        oTitleStyle.alignment = alignment_center


        font = xlwt.Font() # Create Font
        font.colour_index = 4 # 蓝色字体
        font.underline=True
        oLinkStyle = xlwt.XFStyle()
        oLinkStyle.font = font

        sheet = oBook.add_sheet("sheet", True)
        for iRow, tInfo in enumerate(lstAnswer):
            for iCol, text in enumerate(tInfo):
                if iRow == 0:
                    sheet.write(iRow, iCol, str(text), oTitleStyle)
                elif iCol == 5:
                    sheet.write(iRow, iCol, text, oLinkStyle)
                else:
                    sheet.write(iRow, iCol, str(text))

        sheet.col(0).width = 756 * 20
        sheet.col(1).width = 256 * 20
        sheet.col(5).width = 2256 * 20
        oBook.save(sFileName)


    async def Parse(self, url, dInfo, html):
        iType = dInfo["priority"]
        soup = BeautifulSoup(html, 'lxml')
        misc.Write2File("html", html, "w")
        if iType == 0:
            await self.ParsePage(url, dInfo, soup)

    def _Replace(self, msg):
        for txt in self.m_DelList:
            msg = msg.replace(txt, "")
        return msg

    async def ParsePage(self, pageurl, dPageInfo, soup):
        iNum = 0
        for oDivItem in soup.findAll("div", {"class":"item"}):
            oA = oDivItem.find("a", {"class":re.compile("product*")})
            itemurl = oA.get("href").replace("//", "")
            iIndex = itemurl.find("?")
            itemurl = itemurl[:iIndex]

            name = oA.text
            oPrice = oDivItem.find("span", {"itemprop":"price"})
            price = oPrice.text
            oRateNum = oDivItem.find("a", {"class":re.compile("rate.*num*")})
            if oRateNum:
                feedback = oRateNum.text
            else:
                feedback = "0"
            iFeedBack = int(self._Replace(feedback))
            oTotalOrders = oDivItem.find("em", {"title":"Total Orders"})
            orders = oTotalOrders.text
            iOrder = int(self._Replace(orders))
            dItemInfo = {
                "priority"  :2,
                "parent"    :pageurl,
                "time"      :misc.GetSecond(),
                "name"      :name,
                "price"     :price,
                "feedback"  :iFeedBack,
                "order"     :iOrder,
                "xxoo"      :iOrder - iFeedBack,
            }
            self.m_DoneInfo[itemurl] = dItemInfo
            print("%s | %s | %s | %s" % (name, price, iFeedBack, iOrder))
            iNum += 1

        dPageInfo = self.m_DoingUrl.pop(pageurl)
        if iNum:
            self.m_DoneInfo[pageurl] = dPageInfo
        else:
            print("error %s" % pageurl)
