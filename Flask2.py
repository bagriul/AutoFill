from flask import Flask, request, make_response
from flask_cors import CORS
import warnings
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import dropbox
from dropbox.exceptions import ApiError
import json
import requests
import browser_cookie3
from lxml import html
from urllib3.exceptions import InsecureRequestWarning

app = Flask(__name__)
dbx = dropbox.Dropbox('t7JS5c0ogfcAAAAAAAAAAeyldeVzBN_BPDVBHiSoxxpTro9pTHIqsTBMHaouKpba')
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings('ignore', category=InsecureRequestWarning)
CORS(app, resources={r'/*': {'origins': '*'}}, supports_credentials=True)


def Zapovnenya():
    global url
    def Benef():
        cookies = browser_cookie3.chrome(
            cookie_file='\GalContract+Vkursi Cookies')
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36",
            "Accept-Encoding": "gzip, deflate",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "DNT": "1",
            "Connection": "close",
            "Upgrade-Insecure-Requests": "1"
        }
        global response
        response = requests.get('https://vkursi.pro/card/' + code, verify=False, headers=headers, cookies=cookies)

        edrpou = code

        html = response.text

        splited = html.splitlines()
        for benef in splited:
            if "let companyBeneficiars" in benef:
                companyBeneficiars = benef[40:-3]

        newCB = companyBeneficiars.split('},{')
        n = -1
        lenCB = len(newCB)
        beneficiarsList = []
        while n < lenCB:
            try:
                n += 1
                element = newCB[n]
                data = "{" + element + "}"
                data_json = json.loads(data)
                try:
                    line = data_json['personName'] + ", який(-а) проживає за адресою: " + data_json['address']
                    beneficiarsList.append(line)
                except (TypeError, KeyError) as e:
                    try:
                        line = data_json['personName'] + ", який(-а) проживає за адресою: " + "АДРЕСУ НЕ ЗНАЙДЕНО!"
                        beneficiarsList.append(line)
                    except (TypeError, KeyError) as e:
                        typeofr = input(
                            "Причина відсутності кінцевих бенефіціарів:\n1 - жодна особа не підпадає під критерії поняття кінцевого бенефіціарного власника\n2 - кп\n3 - інше\n")
                        typeofreason = str(typeofr)
                        global reason
                        if typeofreason == "1":
                            reason = "жодна особа не підпадає під критерії поняття кінцевого бенефіціарного власника відповідно до законодавства України"
                        if typeofreason == "2":
                            reason = "відповідно до ч.9 ст.9 Закону України «Про державну реєстрацію  юридичних осіб, фізичних осіб-підприємців та громадських формувань» до Єдиного державного реєстру юридичних осіб, фізичних осіб - підприємців та громадських формувань не вноситься інформація про кінцевого бенефіціарного власника комунальних підприємств".decode(
                                'utf-8')
                        if typeofreason == "3":
                            reason = input("Причина: ")
                        beneficiarsList.append(reason)
            except IndexError:
                break

        df = pd.DataFrame(beneficiarsList)
        global df2
        df2 = df.to_string(index=False, header=False)

    if "LLP" in str(auctionID):
        OrendaOrPryvOrZem = "1"
    elif "LLE" in str(auctionID):
        OrendaOrPryvOrZem = "1"
    elif "LLD" in str(auctionID):
        OrendaOrPryvOrZem = "1"
    elif "UA-PS" in str(auctionID):
        OrendaOrPryvOrZem = "2"
    elif "LRE" in str(auctionID):
        OrendaOrPryvOrZem = "3"
    elif "LSP" in str(auctionID):
        OrendaOrPryvOrZem = "3"
    elif "LSE" in str(auctionID):
        OrendaOrPryvOrZem = "3"

    member = code
    auctionLink = auctionID

    # auctionLink = "https://prozorro.sale/auction/LLE001-UA-20211231-37299"
    # member = "Ханас Роман"

    cookies = browser_cookie3.chrome(domain_name='.tsbgalcontract.org.ua',
                                     cookie_file='\GalContract+Vkursi Cookies')
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36",
        "Accept-Encoding": "gzip, deflate",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "DNT": "1",
        "Connection": "close",
        "Upgrade-Insecure-Requests": "1"
    }
    response = requests.get(
        'https://sales.tsbgalcontract.org.ua/EditDataHandler.ashx?CN=0&CommandName=GetMembers&page=1&rows=1000000&sidx=id&sord=desc&TimeMark=80836674',
        verify=False, headers=headers, cookies=cookies)
    AllUsers = response.text

    def UserID(AllUsers, member):
        for UserID in json.loads(AllUsers)['rows']:
            if UserID['full_name'].__contains__(member):
                return UserID['id']
            elif UserID['short_name'].__contains__(member):
                return UserID['id']
            elif UserID['code'].__contains__(member):
                return UserID['id']

    MemberID = UserID(AllUsers, member)

    UserInfoPage = "https://sales.tsbgalcontract.org.ua/DataHandler.ashx?CN=0&CommandName=jGetDetailMember&id=" + MemberID
    cookies = browser_cookie3.chrome(domain_name='.tsbgalcontract.org.ua',
                                     cookie_file='\GalContract+Vkursi Cookies')
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36",
        "Accept-Encoding": "gzip, deflate",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "DNT": "1",
        "Connection": "close",
        "Upgrade-Insecure-Requests": "1"
    }
    response = requests.get(UserInfoPage, verify=False, headers=headers, cookies=cookies)
    UserInfoText = response.text
    UserInfoJson = json.loads(UserInfoText)['member']
    # IBAN = UserInfoJson['iban']
    FullName1 = UserInfoJson['short_name']
    PostalCode = UserInfoJson['postalCode']
    Locality = UserInfoJson['locality']
    StreetAddress = UserInfoJson['streetAddress']
    Adresa = PostalCode + ", " + Locality + ", " + StreetAddress
    Code = UserInfoJson['code']
    mainPP_position = UserInfoJson['mainPP_position']
    mainPP_name = UserInfoJson['mainPP_name']
    ShortName = UserInfoJson['short_name']
    member_type = UserInfoJson['member_type']

    if str(member_type) == "904":
        typeofdoc = "1"
    elif "ФОП" in str(UserInfoJson):
        typeofdoc = "3"
    else:
        typeofdoc = "2"

    if typeofdoc == "2":
        NaPidstavi = "Статуту"

    f = datetime.now()
    day = str(f.day)
    if len(day) == 1:
        day2 = "0" + day
    elif len(day) == 2:
        day2 = day
    month = str(f.month)
    year = str(f.year)
    for old, new in [('12', 'грудня'), ('1', 'січня'), ('11', 'листопада'), ('2', 'лютого'), ('3', 'березня'),
                     ('4', 'квітня'), ('5', 'травня'), ('6', 'червня'), ('7', 'липня'), ('8', 'серпня'),
                     ('9', 'вересня'), ('10', 'жовтня')]:
        month = month.replace(old, new)

    if (OrendaOrPryvOrZem == "1") or (OrendaOrPryvOrZem =="3"):
        if auctionLink.__contains__("prozorro.sale"):
            page = requests.get(auctionLink)
        elif auctionLink.__contains__("tsbgalcontract"):
            numForLink = auctionLink[44:]
            page = requests.get("https://prozorro.sale/auction/" + numForLink)
        else:
            page = requests.get("https://prozorro.sale/auction/" + auctionLink)
        tree = html.fromstring(page.content)
        hashID = tree.xpath('//*[@id="__next"]/main/header/div/div/div[1]/p/text()')
        page = requests.get('https://procedure.prozorro.sale/api/procedures/' + hashID[0][27:])
        data_json = page.json()
        LotTitle = data_json['title']['uk_UA']
        if auctionLink.__contains__("prozorro.sale"):
            LotNumber = auctionLink[30:]
        elif auctionLink.__contains__("tsbgalcontract"):
            LotNumber = auctionLink[44:-1]
        else:
            LotNumber = auctionLink
        if LotTitle.__contains__("адресою"):
            LotTitle = LotTitle
        else:
            LotAdresa = data_json['items'][0]['address']['locality']['uk_UA'] + \
                        data_json['items'][0]['address']['streetAddress']['uk_UA']
            LotTitle = str(LotTitle) + ",за адресою: " + str(LotAdresa)
    elif OrendaOrPryvOrZem == "2":
        if auctionLink.__contains__("prozorro.sale"):
            page = requests.get(auctionLink)
        elif auctionLink.__contains__("tsbgalcontract"):
            numForLink = auctionLink[44:]
            page = requests.get("https://prozorro.sale/auction/" + numForLink)
        else:
            page = requests.get("https://prozorro.sale/auction/" + auctionLink)
        tree = html.fromstring(page.content)
        hashID = tree.xpath('//*[@id="__next"]/main/header/div/div/div[1]/p/text()')
        page = requests.get('https://public.api.ea2.openprocurement.net/api/0/auctions/' + hashID[0][28:])
        data_json = page.json()
        LotTitle = data_json['data']['title']
        if auctionLink.__contains__("prozorro.sale"):
            LotNumber = auctionLink[30:]
        elif auctionLink.__contains__("tsbgalcontract"):
            LotNumber = auctionLink[44:-1]
        else:
            LotNumber = auctionLink
        if LotTitle.__contains__("адресою"):
            LotTitle = LotTitle
        else:
            LotAdresa = data_json['items'][0]['address']['locality'] + data_json['items'][0]['address']['streetAddress']
            LotTitle = str(LotTitle) + ",за адресою: " + str(LotAdresa)

    if typeofdoc == "1":
        list = FullName1.split(" ")
        surname = list[0]
        pib1 = list[1][0]
        pib2 = list[2][0]
        pibfull = surname + " " + pib1 + ". " + pib2 + "."
    elif typeofdoc == "2":
        mainPP_name1 = mainPP_name
        list = mainPP_name1.split(" ")
        surname = list[0]
        pib1 = list[1][0]
        pib2 = list[2][0]
        mainPP_pibfull = surname + " " + pib1 + ". " + pib2 + "."
        Benef()
    elif typeofdoc == "3":
        FullName2 = FullName1[4:]
        list = FullName2.split(" ")
        surname = list[0]
        pib1 = list[1][0]
        pib2 = list[2][0]
        pibfull = surname + " " + pib1 + ". " + pib2 + "."

    print(Code)
    print(Adresa)
    print(FullName1)
    # print(pibfull)
    print(month)
    print(day2)

    if typeofdoc == "1":
        if str(OrendaOrPryvOrZem) == "1":
            doc = DocxTemplate("\Шаблони\Шаблони_оренда_ФІЗ\Зразок_заяви_на_участь_для_ФІЗ.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                    'LotNumber': LotNumber,
                    'day2': day2, 'month': month, 'pibfull': pibfull, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь_ФІЗ" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
                with open(filename, "rb") as f:
                    dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
                shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
                url_dl0 = str(shared_link_metadata.url)
                url = url_dl0.replace('dl=0', 'dl=1')
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
                with open(filename, "rb") as f:
                    dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
                shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
                url_dl0 = str(shared_link_metadata.url)
                url = url_dl0.replace('dl=0', 'dl=1')
        elif str(OrendaOrPryvOrZem) == "2":
            doc = DocxTemplate("Шаблони\Шаблони_приватизація_ФІЗ\Заява_на_участь_ФІЗ_приватизація.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'pibfull': pibfull, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь_ФІЗ_приватизація" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)

            doc = DocxTemplate("Шаблони\Шаблони_приватизація_ФІЗ\Згода_з_умовами_ФІЗ_приватизація.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'pibfull': pibfull, 'year': year}
            doc.render(context)
            filename = "Згода_з_умовами_ФІЗ_приватизація" + Code + ".docx"
            doc.save(filename)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
            shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
            url_dl0 = str(shared_link_metadata.url)
            url = url_dl0.replace('dl=0', 'dl=1')

        elif str(OrendaOrPryvOrZem) == "3":
            doc = DocxTemplate("Шаблони\Шаблони_земля_ФІЗ\Зразок_заяви_на_участь_ЗЕМЛЯ_для_ФІЗ.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'pibfull': pibfull, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь_земля_ФІЗ" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
                with open(filename, "rb") as f:
                    dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
                shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
                url_dl0 = str(shared_link_metadata.url)
                url = url_dl0.replace('dl=0', 'dl=1')
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
                with open(filename, "rb") as f:
                    dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
                shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
                url_dl0 = str(shared_link_metadata.url)
                url = url_dl0.replace('dl=0', 'dl=1')

    elif typeofdoc == "3":
        if str(OrendaOrPryvOrZem) == '1':
            doc = DocxTemplate("Шаблони\Шаблони_оренда_ФОП\Зразок_заяви_на_участь_для_ФОП.docx")
            context = {'Code': Code, 'FullName2': FullName2, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'pibfull': pibfull, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь_оренда_ФОП" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
                with open(filename, "rb") as f:
                    dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
                shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
                url_dl0 = str(shared_link_metadata.url)
                url = url_dl0.replace('dl=0', 'dl=1')
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
                with open(filename, "rb") as f:
                    dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
                shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
                url_dl0 = str(shared_link_metadata.url)
                url = url_dl0.replace('dl=0', 'dl=1')
        elif str(OrendaOrPryvOrZem) == '2':
            doc = DocxTemplate("Шаблони\Шаблони_приватизація_ФОП\Заява_на_участь_ФОП_приватизація.docx")
            context = {'Code': Code, 'FullName2': FullName2, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'pibfull': pibfull, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь_ФОП_приватизація" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)

            doc = DocxTemplate("Шаблони\Шаблони_приватизація_ФОП\Згода_з_умовами_ФОП_приватизація.docx")
            context = {'Code': Code, 'FullName2': FullName2, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'pibfull': pibfull, 'year': year}
            doc.render(context)
            filename = "Згода_з_умовами_ФОП_приватизація" + Code + ".docx"
            doc.save(filename)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
            shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
            url_dl0 = str(shared_link_metadata.url)
            url = url_dl0.replace('dl=0', 'dl=1')
        elif str(OrendaOrPryvOrZem) == "3":
            doc = DocxTemplate("Шаблони\Шаблони_земля_ФОП\Зразок_заяви_на_участь_ЗЕМЛЯ_для_ФОП.docx")
            context = {'Code': Code, 'FullName2': FullName2, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'pibfull': pibfull, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь_земля_ФОП" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
                with open(filename, "rb") as f:
                    dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
                shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
                url_dl0 = str(shared_link_metadata.url)
                url = url_dl0.replace('dl=0', 'dl=1')
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
                with open(filename, "rb") as f:
                    dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
                shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
                url_dl0 = str(shared_link_metadata.url)
                url = url_dl0.replace('dl=0', 'dl=1')

    elif typeofdoc == "2":
        if str(OrendaOrPryvOrZem) == '1':
            doc = DocxTemplate("Шаблони\Шаблони_оренда_ТЗОВ\Зразок_заяви_на_участь_для_ЮО.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь_ЮО" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)

            doc = DocxTemplate("Шаблони\Шаблони_оренда_ТЗОВ\Форма_Довідки_про_бенефіціарних_власників.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'Benef': df2, 'year': year}
            doc.render(context)
            filename = "Форма_Довідки_про_бенефіціарних_власників_r" + Code + ".docx"
            doc.save(filename)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
            shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
            url_dl0 = str(shared_link_metadata.url)
            url = url_dl0.replace('dl=0', 'dl=1')

            doc = DocxTemplate("Довідка_про_відсутність_кінцевого_бенефіциарного_власника.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'reason': df2, 'year': year}
            doc.render(context)
            doc.save("Довідка_про_відсутність_кінцевого_бенефіциарного_власника_r.docx")
        elif str(OrendaOrPryvOrZem) == '2':
            doc = DocxTemplate("Шаблони\Шаблони_приватизація_ТЗОВ\Форма_Довідки_про_бенефіціарних_власників.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'Benef': df2, 'year': year}
            doc.render(context)
            filename = "Форма_Довідки_про_бенефіціарних_власників" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)

            doc = DocxTemplate("Довідка_про_відсутність_кінцевого_бенефіциарного_власника.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'reason': df2, 'year': year}
            doc.render(context)
            doc.save("Довідка_про_відсутність_кінцевого_бенефіциарного_власника_r.docx")

            doc = DocxTemplate("Шаблони\Шаблони_приватизація_ТЗОВ\Заява_на_участь(приватизація).docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'reason': df2, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь(приватизація)_r" + Code + ".docx"
            doc.save(filename)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)

            doc = DocxTemplate("Шаблони\Шаблони_приватизація_ТЗОВ\Згода_з_умовами_аукціону(приватизація).docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'reason': df2, 'year': year}
            doc.render(context)
            filename = "Згода_з_умовами_аукціону(приватизація)_r" + Code + ".docx"
            doc.save(filename)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
            shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
            url_dl0 = str(shared_link_metadata.url)
            url = url_dl0.replace('dl=0', 'dl=1')
        elif str(OrendaOrPryvOrZem) == "3":
            doc = DocxTemplate("Шаблони\Шаблони_земля_ТЗОВ\Зразок_заяви_на_участь_ЗЕМЛЯ_для_ЮО.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'reason': df2, 'year': year}
            doc.render(context)
            filename = "Заява_на_участь_ЗЕМЛЯ_r" + Code + ".docx"
            doc.save(filename)
            try:
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            except ApiError:
                dbx.files_delete_v2('/folder_' + Code)
                folder = dbx.files_create_folder_v2('/folder_' + Code)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)

            doc = DocxTemplate("Шаблони\Шаблони_земля_ТЗОВ\Форма_Довідки_про_бенефіціарних_власників.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'Benef': df2, 'year': year}
            doc.render(context)
            filename = "Форма_Довідки_про_бенефіціарних_власників_ЗЕМЛЯ_r" + Code + ".docx"
            doc.save(filename)
            with open(filename, "rb") as f:
                dbx.files_upload(f.read(), '/folder_' + str(Code) + '/' + filename, mute=True)
            shared_link_metadata = dbx.sharing_create_shared_link('/folder_' + Code)
            url_dl0 = str(shared_link_metadata.url)
            url = url_dl0.replace('dl=0', 'dl=1')

            doc = DocxTemplate("Довідка_про_відсутність_кінцевого_бенефіциарного_власника.docx")
            context = {'Code': Code, 'FullName1': FullName1, 'Adresa': Adresa, 'LotTitle': LotTitle,
                       'LotNumber': LotNumber,
                       'day2': day2, 'month': month, 'mainPP_position': mainPP_position, 'mainPP_name': mainPP_name,
                       'NaPidstavi': NaPidstavi, 'ShortName': ShortName, 'mainPP_pibfull': mainPP_pibfull,
                       'reason': df2, 'year': year}
            doc.render(context)
            doc.save("Довідка_про_відсутність_кінцевого_бенефіциарного_власника_r.docx")

@app.route('/calculate', methods=['GET', 'POST'])
def calculator():
    data = request.get_data().decode('utf-8')
    datas = str(data).split(",")
    global code
    if datas[0] == "":
        code = datas[1]
    else:
        code = datas[0]
    global auctionID
    auctionID = datas[2]

    Zapovnenya()
    print(url)

    return url

if __name__ == '__main__':
    app.run(threaded=True)
