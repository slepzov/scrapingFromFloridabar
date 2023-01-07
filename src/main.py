import requests
from bs4 import BeautifulSoup
from time import sleep
import xlsxwriter
from random import choice


desktop_agents = [

    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_1) AppleWebKit/602.2.14 (KHTML, like Gecko) Version/10.0.1 '
    'Safari/602.2.14',
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.71 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.98 '
    'Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.98 '
    'Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.71 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:50.0) Gecko/20100101 Firefox/50.0']


def random_headers():
    return {'User-Agent': choice(desktop_agents),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'}


def decode(g):
    r = int(g[:2],16)
    email = ''.join([chr(int(g[i:i+2], 16) ^ r) for i in range(2, len(g), 2)])
    return email


def get_info():
    for count in range(1, 6):

        url = f"https://www.floridabar.org/directories/find-mbr/?sdx=N&eligible=N&deceased=N&pracAreas=C18&pageNumber={count}&pageSize=50"
        response = requests.get(url, headers=random_headers())
        soup = BeautifulSoup(response.text, "lxml")
        data = soup.findAll("li", class_="profile-compact")

        for i in data:
            email = decode(i.find("div", class_="profile-contact").find("a", class_="icon-email").get("href").split("#")[-1])
            mail_address = i.find("div", class_="profile-contact").find("p").get_text(strip=True, separator='\n').splitlines()
            if len(mail_address) < 3:
                company = "None"
                address = mail_address[0]
                city_address = mail_address[1].split(",")
            elif len(mail_address) > 3:
                company = mail_address[0]
                address = mail_address[1] + mail_address[2]
                city_address = mail_address[3].split(",")
            else:
                company = mail_address[0]
                address = mail_address[1]
                city_address = mail_address[2].split(",")
            city = city_address[0]
            state = city_address[1].strip().split(" ")[0]
            zip_code = city_address[1].strip().split(" ")[1]
            phone = i.find("div", class_="profile-contact").find("a").text
            website = email.split("@")[1]
            lo2_bar = i.find("p", class_="profile-bar-number").find("span").text[1:]
            full_name = i.find("p", class_="profile-name").text.split(" ")
            first_name = full_name[0].strip()
            last_name = full_name[-1].strip()

            sleep(1)
            yield email, company, address, city, state, zip_code, phone, website, lo2_bar, first_name, last_name


def writer(parametr):
    book = xlsxwriter.Workbook(r"C:\data.xlsx")
    page = book.add_worksheet("product")

    row = 1
    column = 0

    page.set_column("A:A", 30)
    page.set_column("B:B", 25)
    page.set_column("C:C", 30)
    page.set_column("D:D", 15)
    page.set_column("E:E", 10)
    page.set_column("F:F", 15)
    page.set_column("G:G", 15)
    page.set_column("H:H", 20)
    page.set_column("I:I", 10)
    page.set_column("J:J", 10)
    page.set_column("K:K", 10)

    count = 1
    for item in parametr:
        page.write(row, column, item[0])
        page.write(row, column + 1, item[1])
        page.write(row, column + 2, item[2])
        page.write(row, column + 3, item[3])
        page.write(row, column + 4, item[4])
        page.write(row, column + 5, item[5])
        page.write(row, column + 6, item[6])
        page.write(row, column + 7, item[7])
        page.write(row, column + 8, item[8])
        page.write(row, column + 9, item[9])
        page.write(row, column + 10, item[10])
        print(f"Внесено записей: {count}")
        row += 1
        count += 1

    book.close()


def main():
    writer(get_info())


if __name__ == "__main__":
    main()

