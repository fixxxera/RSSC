import datetime
import json
import sqlite3

import os
import requests
import xlsxwriter
from bs4 import BeautifulSoup
from multiprocessing.dummy import Pool as ThreadPool

session = requests.session()
def get_proxy():
    print("Looking for a working proxy server")
    soup = BeautifulSoup(requests.get('https://www.us-proxy.org').text, 'lxml')
    table = soup.find('table', {'id': 'proxylisttable'})
    tbody = table.find('tbody')
    proxies = []
    for tr in tbody:
        columns = tr.find_all('td')
        if columns[2].text in 'US' and columns[4].text in 'anonymous' and columns[6].text in 'yes':
            proxies.append("https://" + columns[0].text + ":" + columns[1].text)
    for p in proxies:
        try:
            proxy_line = {'https': p}
            resp = requests.get('https://www.rssc.com/cruises', proxies=proxy_line, timeout=10)
            if resp.ok:
                print("Found one!")
                return proxy_line
            else:
                print(p, "Not working")
        except requests.exceptions.ProxyError:
            print(p, "Not working")
        except requests.exceptions.ConnectTimeout:
            print(p, "Not working")
        except requests.exceptions.ReadTimeout:
            print(p, "Not working")
proxies = get_proxy()
page = session.get('https://www.rssc.com/cruises', proxies=proxies)
soup = BeautifulSoup(page.text, 'lxml')
voyages = []
result_list = []


pool = ThreadPool(4)


def convert_date(day, month, year):
    if month == 'Jan':
        month = '1'
    elif month == 'Feb':
        month = '2'
    elif month == 'Mar':
        month = '3'
    elif month == 'Apr':
        month = '4'
    elif month == 'May':
        month = '5'
    elif month == 'Jun':
        month = '6'
    elif month == 'Jul':
        month = '7'
    elif month == 'Aug':
        month = '8'
    elif month == 'Sep':
        month = '9'
    elif month == 'Oct':
        month = '10'
    elif month == 'Nov':
        month = '11'
    elif month == 'Dec':
        month = '12'
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def calculate_days(date, duration):
    dateobj = datetime.datetime.strptime(date, "%m/%d/%Y")
    calculated = dateobj + datetime.timedelta(days=int(duration))
    calculated = calculated.strftime("%m/%d/%Y")
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


def split_carib_auto(ports, dc, dn):
    cu = []
    wc = []
    ec = []
    bm = []
    conn = sqlite3.connect(r'/home/fixxxer/PycharmProjects/PortsExplorer/ports.db')
    c = conn.cursor()
    c.execute("SELECT * FROM portlist WHERE destination_name='Cuba'")
    for row in c.fetchall():
        cu.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='West Carib'")
    for row in c.fetchall():
        wc.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='East Carib'")
    for row in c.fetchall():
        ec.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='Bermuda'")
    for row in c.fetchall():
        bm.append(row[0])
    c.close()
    conn.close()
    result = []
    iscu = False
    isec = False
    iswc = False
    ports_list = []
    for i in range(len(ports)):
        if i == 0:
            pass
        else:
            ports_list.append(ports[i])
    for element in cu:
        for p in ports_list:
            if p in element or element in p:
                iscu = True
    if not iscu:
        for element in wc:
            for p in ports_list:
                if p in element or element in p:
                    iswc = True
    if not iswc:
        for element in ec:
            for p in ports_list:
                if p in element or element in p:
                    isec = True
    if iscu:
        result.append("Cuba")
        result.append("C")
        result.append("CU")
        return result
    elif iswc:
        result.append("West Carib")
        result.append("C")
        result.append("WC")
        return result
    elif isec:
        result.append("East Carib")
        result.append("C")
        result.append("EC")
        return result
    else:
        result.append(dn)
        result.append(dc)
        result.append("")
        return result


def split_europe_auto(ports, dn, dc):
    baltic = []
    eastern_med = []
    west_med = []
    baltic.append("SOU")
    baltic.append("HAU")
    baltic.append("FLM")
    baltic.append("AES")
    baltic.append("BGO")
    baltic.append("KWL")
    baltic.append("GNR")
    baltic.append("SVG")
    baltic.append("TOS")
    baltic.append("LKN")
    baltic.append("HVG")
    conn = sqlite3.connect(r'/home/fixxxer/PycharmProjects/PortsExplorer/ports.db')
    c = conn.cursor()
    c.execute("SELECT * FROM portlist WHERE destination_name='Baltics'")
    for row in c.fetchall():
        baltic.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='EastMed'")
    for row in c.fetchall():
        eastern_med.append(row[0])
    c.execute("SELECT * FROM portlist WHERE destination_name='WestMed'")
    for row in c.fetchall():
        west_med.append(row[0])
    c.close()
    conn.close()
    for element in baltic:
        for p in ports:
            if p in element or element in p:
                return ['Baltic', 'E']
            elif ports[0] in element or element in ports[0]:
                return ['Baltic', 'E']

    for element in eastern_med:
        for p in ports:
            if p in element or element in p:
                return ['Eastern Med', 'E']

    for element in west_med:
        for p in ports:
            if p in element or element in p:
                return ['Western Med', 'E']
    return [dn, dc]


scripts = soup.find_all('script')
for script in scripts:
    if 'cruises = ' in script.text:
        sr = script.text.split(' = ')[1].strip()[:-1]
        jsonfile = (json.loads(sr))
        for j in jsonfile:
            voyages.append(j)


def parse(v):
    ports = []
    vessel_name = v['ship']['title']
    ship_id = get_from_vessel_name(vessel_name)
    destination = get_from_code(v['region']['id'])
    destination_code = destination[1]
    destination_name = destination[0]
    cruise_id = '13'
    itinerary_id = ''
    cruise_line_name = 'RSSC'
    vessel_id = v['voyageId']
    for port in v['ports']:
        ports.append(port['title'])
    number_of_nights = v['duration']
    trip_page = 'https://www.rssc.com/cruises/' + vessel_id + '/summary'
    print("Downloading", trip_page)
    response = requests.get(trip_page, proxies=proxies)
    detail_soup = BeautifulSoup(response.text, 'lxml')
    brochure_name = detail_soup.find('span', {'class': 'headline-first'}).text.strip()
    h1 = detail_soup.find('h1', {'class': 'headline a1m'})
    date = h1.text.split(' | ')[2].replace('Departs ', '').replace(',', '').split()
    sail_date = convert_date(date[1], date[0], date[2])
    return_date = calculate_days(sail_date, number_of_nights)
    prices = []
    if destination_name == 'Caribbean/Panama Canal':
        destination = split_carib_auto(ports, destination_code, destination_name)
        destination_name = destination[0]
        destination_code = destination[1]
    if 'WMED/EMED' in destination_name:
        dest = split_europe_auto(ports, destination_name, destination_code)
        destination_code = dest[1]
        destination_name = dest[0]
    price_table = detail_soup.find('div', {'class', 'fares-table'})
    trs = price_table.find_all('tr', {'class': 'js-toggle-dropdown'})
    for tr in trs:
        room_type = str(tr.find_all('td')[1].find('a').text).strip().split(' Suite ')[0].strip()
        price = str(tr.find_all('td')[3].find_all('span', {'class': 'data-info'})[1].text).replace('$', '').replace(',',
                                                                                                                    '')
        prices.append([room_type, price])
    prices = list(reversed(prices))
    interior_bucket_price = ''
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
    temp = [destination_code, destination_name, ship_id, vessel_name, cruise_id, cruise_line_name,
            itinerary_id,
            brochure_name, number_of_nights, sail_date, return_date,
            interior_bucket_price, oceanview_bucket_price, balcony_bucket_price, suite_bucket_price, ports]
    tmp2 = [temp]
    print(temp)
    result_list.append(tmp2)


pool.map(parse, voyages)
pool.close()
pool.join()


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
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
