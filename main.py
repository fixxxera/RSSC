import datetime
import math
import os

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from multiprocessing.dummy import Pool as ThreadPool

session = requests.session()

# proxies = {'https': 'https://165.138.65.233:3128'}
# proxies = {'https': 'https://192.241.145.201:8080'}
# proxies = {'https': 'https://104.198.223.14:80'}
# proxies = {'https': 'https://35.185.23.159:80'}
proxies = {'https': 'https://70.35.197.74:80'}
pool = ThreadPool(4)
destination_list = ['AFRIND', 'ALSKA', 'ASIAS', 'CANNE', 'CARMX', 'GRNDX', 'EURMD', 'RUSBA', 'LATAM', 'GRNDV']
to_walk = []
result_list = []
total_results = 0


def convert_date(unformated):
    splitter = unformated.split()
    day = splitter[1]
    month = splitter[0]
    year = splitter[2]
    if month == 'January':
        month = '1'
    elif month == 'February':
        month = '2'
    elif month == 'March':
        month = '3'
    elif month == 'April':
        month = '4'
    elif month == 'May':
        month = '5'
    elif month == 'June':
        month = '6'
    elif month == 'July':
        month = '7'
    elif month == 'August':
        month = '8'
    elif month == 'September':
        month = '9'
    elif month == 'October':
        month = '10'
    elif month == 'November':
        month = '11'
    elif month == 'December':
        month = '12'
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def get_total_results(de):
    url = "https://www.rssc.com/cruises/default.aspx?m=&r=" + de + "&dy=&sh=&p=&sp="
    headers = {
        'Host': 'www.rssc.com',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:53.0) Gecko/20100101 Firefox/53.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate, br',
        'Referer': 'https://www.rssc.com/',
        'Cookie': 'AKCountry=US; utag_main=v_id:015a3cd6f81f0015e7883104dd6f01044001500900bd0$_sn:1$_ss:0$_pn:4%3Bexp'
                  '-session$_st:1487081459123$ses_id:1487079405599%3Bexp-session; s_cc=true; '
                  's_sq=rsscprod%3D%2526c.%2526a.%2526activitymap.%2526page%253DBooking%252520funnel%25253A'
                  '%252520Step%2525200%25253A%252520Search%252520results%2526link%253D1%2526region%253Dpagination'
                  '%2526pageIDType%253D1%2526.activitymap%2526.a%2526.c%2526pid%253DBooking%252520funnel%25253A'
                  '%252520Step%2525200%25253A%252520Search%252520results%2526pidt%253D1%2526oid%253Djavascript%25253A'
                  '%25252520CruiseResultsItemClick%252528%252527page%252527%25252C%25252520%2525271%252527%25252C'
                  '%25252520%252527%252527%25252C%25252520%252527%252527%252529%2526ot%253DA%26bgtrsscprod%3D%2526pid'
                  '%253Drssc%25257Ccruises%25257Cna%25257Cna%25257Cna%25257Cna%25257Cna%2526pidt%253D1%2526oid'
                  '%253Djavascript%25253A%25252520CruiseResultsItemClick%252528%252527page%252527%25252C%25252520'
                  '%2525271%252527%25252C%25252520%252527%252527%25252C%25252520%252527%252527%252529%2526ot%253DA; '
                  's_getNewRepeat=1487079659816-New; s_fid=7903A570C377CC63-29C6DF9A1D314E4B; '
                  '_ga=GA1.2.1688205592.1487079406; _gat_tealium_0=1; s_vi=[CS]v1|2C5183F705014C8E-6000014860004F6C['
                  'CE]; ipe_s=a2d3df6d-f916-9020-bc5c-014f15010f14; ipe.370.pageViewedCount=2; '
                  'UserSettings=Country=US; CruiseSearch.GalleryView=True; IPE125021=IPE125021',
        'Proxy-Authorization': 'Basic bWFydGluLmJhbHR1aGluQGdtYWlsLmNvbTphLzZFZzN2T3J3blpiUWxNcWYrdVF5VzNyOU5OU0tTNw==',
        'X-Time': '1487079696226',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Cache-Control': 'max-age=0'
    }
    page = session.get(url=url, headers=headers, proxies=proxies).text
    soup = BeautifulSoup(page, "lxml")
    div = soup.find('div', {'id': 'matchInfo'})
    number = int(div.find('h3').text.split()[0])
    return number


def calculate_days(date, duration):
    dateobj = datetime.datetime.strptime(date, "%m/%d/%Y")
    calculated = dateobj + datetime.timedelta(days=int(duration))
    calculated = calculated.strftime("%-m/%-d/%Y")
    return calculated


def get_from_code(dc):
    if dc == 'AFRIND':
        return ['Exotics', 'O']
    if dc == 'ALSKA':
        return ['Alaska', 'A']
    if dc == 'ASIAS':
        return ['Asia/Pacific', 'TBD']
    if dc == 'CANNE':
        return ['Canada/New England', 'NN']
    if dc == 'CARMX':
        return ['Caribbean/Panama Canal', 'TBD']
    if dc == 'GRNDX':
        return ['Grand crossings', 'TBD']
    if dc == 'EURMD':
        return ['WMED/EMED', 'TBD']
    if dc == 'RUSBA':
        return ['Baltic', 'E']
    if dc == 'LATAM':
        return ['South America', 'S']
    if dc == 'GRNDV':
        return ['Exotics', 'O']


def get_from_vessel_name(vn):
    if vn == 'Seven Seas Mariner':
        return '106'
    if vn == 'Seven Seas Navigator':
        return '107'
    if vn == 'Seven Seas Voyager':
        return '108'
    if vn == 'Seven Seas Explorer':
        return '693'
    pass


def parse(pack):
    headers = {
        'Host': 'www.rssc.com',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:53.0) Gecko/20100101 Firefox/53.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate, br',
        'Referer': 'https://www.rssc.com/',
        'Cookie': 'AKCountry=US; utag_main=v_id:015a3cd6f81f0015e7883104dd6f01044001500900bd0$_sn:1$_ss:0$_pn:4%3Bexp'
                  '-session$_st:1487081459123$ses_id:1487079405599%3Bexp-session; s_cc=true; '
                  's_sq=rsscprod%3D%2526c.%2526a.%2526activitymap.%2526page%253DBooking%252520funnel%25253A'
                  '%252520Step%2525200%25253A%252520Search%252520results%2526link%253D1%2526region%253Dpagination'
                  '%2526pageIDType%253D1%2526.activitymap%2526.a%2526.c%2526pid%253DBooking%252520funnel%25253A'
                  '%252520Step%2525200%25253A%252520Search%252520results%2526pidt%253D1%2526oid%253Djavascript%25253A'
                  '%25252520CruiseResultsItemClick%252528%252527page%252527%25252C%25252520%2525271%252527%25252C'
                  '%25252520%252527%252527%25252C%25252520%252527%252527%252529%2526ot%253DA%26bgtrsscprod%3D%2526pid'
                  '%253Drssc%25257Ccruises%25257Cna%25257Cna%25257Cna%25257Cna%25257Cna%2526pidt%253D1%2526oid'
                  '%253Djavascript%25253A%25252520CruiseResultsItemClick%252528%252527page%252527%25252C%25252520'
                  '%2525271%252527%25252C%25252520%252527%252527%25252C%25252520%252527%252527%252529%2526ot%253DA; '
                  's_getNewRepeat=1487079659816-New; s_fid=7903A570C377CC63-29C6DF9A1D314E4B; '
                  '_ga=GA1.2.1688205592.1487079406; _gat_tealium_0=1; s_vi=[CS]v1|2C5183F705014C8E-6000014860004F6C['
                  'CE]; ipe_s=a2d3df6d-f916-9020-bc5c-014f15010f14; ipe.370.pageViewedCount=2; '
                  'UserSettings=Country=US; CruiseSearch.GalleryView=True; IPE125021=IPE125021',
        'Proxy-Authorization': 'Basic bWFydGluLmJhbHR1aGluQGdtYWlsLmNvbTphLzZFZzN2T3J3blpiUWxNcWYrdVF5VzNyOU5OU0tTNw==',
        'X-Time': '1487079696226',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'Cache-Control': 'max-age=0'
    }
    count = pack[1]
    des = pack[0]
    while count > 0:
        url = 'https://www.rssc.com/WebServices/CruiseFinder/CruiseFinder.asmx/GetCruiseFinderResults'
        json = {'action': 'page', 'clickedItem': str(count), 'currentSelections': 'R' + des + '',
                'currentPage': str(count),
                'currentView': 'gallery', 'compareCruises': '', 'preSetVoyages': '', 'showTitles': 'True',
                'container': '', 'showRemove': 'False'}
        page = session.post(url=url, headers=headers, json=json, proxies=proxies).json()['d']
        count -= 1
        soup = BeautifulSoup(page, 'lxml')
        results = soup.find_all('div', {'class': 'result'})
        for res in results:
            brochure_name = res.find('div', {'class': 'resultHeader'}).find('a').text
            detail_box = res.find('div', {'class': 'detail'})
            vessel_name = detail_box.find('h4').text
            number_of_nights = detail_box.find('h5').text.split()[0]
            spans = detail_box.find_all('span')
            sail_date = convert_date(spans[0].text.replace(',', ''))
            return_date = calculate_days(sail_date, number_of_nights)
            interior_bucket_price = ''
            destination = get_from_code(des)
            destination_name = destination[0]
            destination_code = destination[1]
            cruise_line_name = 'RSSC'
            itinerary_id = ''
            cruise_id = '13'
            vessel_id = get_from_vessel_name(vessel_name)
            price_block = res.find('div', {'class': 'viewDetail'}).find('a')['href']
            url = 'https://www.rssc.com' + price_block
            page = session.get(url=url, headers=headers, proxies=proxies)
            soup = BeautifulSoup(page.text, 'lxml')
            # getting ports here
            ports_table = soup.find('div', {'id': 'itineraryInfo'})
            rows = ports_table.find_all('tr')
            ports = []
            for row in rows:
                tds = row.find_all('td')
                try:
                    ports.append(tds[2].text.split(',')[0].replace('Cruising the ', '').replace('Cruising ', '').strip())
                except IndexError:
                    pass
            print(ports)
            # getting prices here
            table = soup.find('table', {'id': 'right'})
            row_list = []
            prices = []
            siblings = table.find_all('tr')
            for s in siblings:
                row_list.append(s)
            for index in range(1, len(row_list), 2):
                twoforone = row_list[index].find('td', {'class': 'twoforone'}).text
                room = row_list[index].find('a').text.split(' Suite ')[0]
                prices.append([room, twoforone])
            prices = list(reversed(prices))
            ocview = []
            ver = []
            sui = []
            for index in range(0, len(prices)):
                if prices[index][0] == 'Deluxe Window':
                    ocview.append(prices[index][1])
                elif prices[index][0] == 'Deluxe Veranda' or prices[index][0] == 'Veranda':
                    ver.append(prices[index][1])
                elif prices[index][0] == 'Superior' or prices[index][0] == 'Concierge' or prices[index][0] == 'Penthouse':
                    sui.append(prices[index][1])
            if len(sui) > 0:
                suite_bucket_price = sui[0].replace('$', '').replace(',', '')
            else:
                suite_bucket_price = 'N/A'
            if len(ver) > 0:
                balcony_bucket_price = ver[0].replace('$', '').replace(',', '')
            else:
                balcony_bucket_price = 'N/A'
            if len(ocview) > 0:
                oceanview_bucket_price = ocview[0].replace('$', '').replace(',', '')
            else:
                oceanview_bucket_price = 'N/A'
            temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name,
                    itinerary_id,
                    brochure_name, number_of_nights, sail_date, return_date,
                    interior_bucket_price, oceanview_bucket_price, balcony_bucket_price, suite_bucket_price, ports]
            tmp2 = [temp]
            print(temp)
            result_list.append(tmp2)
    pass


for d in destination_list:
    total = get_total_results(d)
    total_results += total
    if total == 0:
        continue
    elif total <= 12:
        to_walk.append([d, 1])
    else:
        pages = math.ceil(total / 12)
        to_walk.append([d, pages])
pool.map(parse, to_walk)
pool.close()
pool.join()


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    # print(userhome)
    now = datetime.datetime.now()
    path_to_file = userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- RSSC.xlsx'
    if not os.path.exists(userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '/Dropbox/XLSX/For Assia to test/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook(path_to_file)

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 25)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 25)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 30)
    worksheet.set_column("G:G", 20)
    worksheet.set_column("H:H", 50)
    worksheet.set_column("I:I", 20)
    worksheet.set_column("J:J", 20)
    worksheet.set_column("K:K", 20)
    worksheet.set_column("L:L", 20)
    worksheet.set_column("M:M", 25)
    worksheet.set_column("N:N", 20)
    worksheet.set_column("O:O", 20)
    worksheet.set_column("P:P", 100)
    worksheet.write('A1', 'DestinationCode', bold)
    worksheet.write('B1', 'DestinationName', bold)
    worksheet.write('C1', 'VesselID', bold)
    worksheet.write('D1', 'VesselName', bold)
    worksheet.write('E1', 'CruiseID', bold)
    worksheet.write('F1', 'CruiseLineName', bold)
    worksheet.write('G1', 'ItineraryID', bold)
    worksheet.write('H1', 'BrochureName', bold)
    worksheet.write('I1', 'NumberOfNights', bold)
    worksheet.write('J1', 'SailDate', bold)
    worksheet.write('K1', 'ReturnDate', bold)
    worksheet.write('L1', 'InteriorBucketPrice', bold)
    worksheet.write('M1', 'OceanViewBucketPrice', bold)
    worksheet.write('N1', 'BalconyBucketPrice', bold)
    worksheet.write('O1', 'SuiteBucketPrice', bold)
    worksheet.write('P1', 'Ports', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    ordinary_number = workbook.add_format({"num_format": '#,##0'})
    date_format = workbook.add_format({'num_format': 'm d yyyy'})
    centered = workbook.add_format({'bold': True})
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    centered.set_bold(True)
    ordinary_number.set_bold(True)
    ordinary_number.set_align("center")
    date_format.set_align("center")
    centered.set_align("center")
    for ship in data_array:
        for l in ship:
            column_count = 0
            for r in l:
                try:
                    if column_count == 0:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 1:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 2:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 3:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 4:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 5:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 6:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 7:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 8:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 9:
                        date_time = datetime.datetime.strptime(str(r), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, centered)
                        column_count += 1
                    elif column_count == 10:
                        date_time = datetime.datetime.strptime(str(r), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, centered)
                        column_count += 1
                    elif column_count == 11:
                        tmp = str(r)
                        cell = ""
                        number = 0
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 12:
                        tmp = str(r)
                        cell = ""
                        number = 0
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 13:
                        tmp = str(r)
                        cell = ""
                        number = 0
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 14:
                        tmp = str(r)
                        cell = ""
                        number = 0
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 15:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                except ValueError:
                    worksheet.write_string(row_count, column_count, r, centered)
                    column_count += 1
            row_count += 1
    workbook.close()
    pass


write_file_to_excell(result_list)
