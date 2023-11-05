import pyautogui as pg
import pandas as pd
import psutil
import time
import random
import sys
import os
import logging
import logging.config
# import pymsteams

# teams = pymsteams.connectorcard('URL do conector')
logging.config.fileConfig('./logs/logging.conf')
logger = logging.getLogger('root')


class DesktopTools:
    def inicia_software():
        for process in psutil.process_iter(['name']):
            if process.info['name'] == 'Fakturama.exe':
                return True
        return False

    def try_locate_image(imagePath, try_count=0, tries=5):
        while try_count >= 0:
            position = pg.locateOnScreen(imagePath,
                                         grayscale=True,
                                         confidence=0.7)
            time.sleep(1)
            try_count += 1
            logging.debug(try_count)
            if try_count >= tries or position is not None:
                break
        try:
            if position is not None:
                print(f"position = {position}")
                return position
            else:
                raise Exception(f'Imagem: "{imagePath}", não localizada')
        except Exception as error:
            timestr = time.strftime("%Y%m%d-%H%M%S")
            logging.critical(error)
            # teams.text("RPA Fakturama - Erro CRITICO!")
            # teams.send()
            pg.screenshot(f"./logs/screenshots/ERROR_{timestr}.png")
            sys.exit()

    def random_select(archive='./assets/fake_data.csv',):
        if list(archive)[-1] == 'v':
            df = pd.read_csv(archive)
        else:
            df = pd.read_excel(archive)
        rows = (df.shape[0]-1)
        item = str(df.iloc[(random.randint(0, rows)), 0])
        return item


class FakturamaActivities:
    def cadastra_produtos():
        if not DesktopTools.inicia_software():
            os.startfile(r"C:\Program Files\Fakturama2\Fakturama.exe")

        df = pd.read_excel(r".\assets\fakturama.xlsx")
        print(df.head())

        new_product = DesktopTools.try_locate_image(
            r".\assets\images\btn_new_product.PNG", tries=30)

        for i, r in df.iterrows():
            item_number = str(r["Item Number"])
            product_name = r["Name"]
            category = r["Category"]
            gtin = str(r["GTIN"])
            description = r["Description"]
            notice = r["Notice"]
            if new_product is not None:
                pg.click(new_product, interval=2)
                label = DesktopTools.try_locate_image(
                    r".\assets\images\label_new_product.PNG")
                pg.click(label, interval=2)
                pg.press('tab', 2, interval=0.5)
                pg.typewrite(item_number)
                logging.info(item_number)
                pg.press('tab', 1, interval=0.5)
                pg.typewrite(product_name)
                logging.info(product_name)
                pg.press('tab', 1, interval=0.5)
                pg.typewrite(category)
                logging.info(category)
                pg.press('tab', 1, interval=0.5)
                pg.typewrite(gtin)
                logging.info(gtin)
                pg.press('tab', 2, interval=0.5)
                pg.typewrite(description)
                logging.info(description)
                pg.press('tab', 10, interval=0.5)
                pg.typewrite(notice)
                pg.hotkey('ctrl', 's')
                pg.hotkey('ctrl', 'w')
        pg.hotkey('alt', 'F4')

    def cadastra_contato():
        if not DesktopTools.inicia_software():
            os.startfile(r"C:\Program Files\Fakturama2\Fakturama.exe")

        df = pd.read_csv(r".\assets\fake_data.csv")
        print(df.head())

        new_contact = DesktopTools.try_locate_image(
            r".\assets\images\btn_new_contact.PNG", tries=30)

        for i, r in df.iterrows():
            first_name = r.iloc[0]
            last_name = r.iloc[1]
            company = r.iloc[2]
            address = r.iloc[4]
            email = r.iloc[5]
            phone = str(r.iloc[6])
            if new_contact is not None:
                pg.click(new_contact, interval=2)
                pg.press('tab', 2, interval=0.5)
                pg.typewrite(company)
                print(company)
                pg.press('tab', 2, interval=0.5)
                pg.typewrite(first_name)
                print(first_name)
                pg.press('tab')
                pg.typewrite(last_name)
                print(last_name)
                pg.press('tab', 5, interval=0.5)
                pg.typewrite(address)
                print(address)
                pg.press('tab', 7, interval=0.5)
                pg.typewrite(email)
                print(email)
                pg.press('tab')
                pg.typewrite(phone)
                print(phone)

    def preenche_ordem():
        if not DesktopTools.inicia_software():
            os.startfile(r"C:\Program Files\Fakturama2\Fakturama.exe")

        contato = DesktopTools.random_select('./assets/fake_data.csv')

        new_order = DesktopTools.try_locate_image(r".\assets\images\btn_new_order.PNG", tries=30)
        if new_order is not None:
            pg.click(new_order, interval=2)
            contact_list = DesktopTools.try_locate_image(r".\assets\images\btn_contact_list.PNG")
            pg.click(contact_list, interval=2)
            pg.typewrite(contato)
            first_item_select = DesktopTools.try_locate_image(r".\assets\images\btn_first_item.PNG")
            pg.click(first_item_select, interval=2)
            pg.move(40, 0)
            pg.doubleClick()
            for i in range(3):
                produto = DesktopTools.random_select('./assets/fakturama.xlsx')
                product_list = DesktopTools.try_locate_image(r".\assets\images\btn_product_list.PNG")
                pg.click(product_list, interval=2)
                pg.typewrite(produto, interval=0.1)
            pg.hotkey("ctrl", "s")
            pg.hotkey('ctrl', 'w')
            pg.hotkey('alt', 'F4')


if __name__ == '__main__':
    FakturamaActivities.preenche_ordem()
