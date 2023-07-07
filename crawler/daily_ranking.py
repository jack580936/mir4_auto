import csv
import requests

from bs4 import BeautifulSoup
import re
from datetime import datetime
import threading

import sys
from PyQt5.QtCore import Qt,QSettings
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit, QVBoxLayout, QCheckBox, \
    QRadioButton, QGroupBox, QFormLayout, QButtonGroup, QHBoxLayout
from qt_material import apply_stylesheet

png_to_class = {
    'char_1': '戰士',
    'char_2': '術士',
    'char_3': '道士',
    'char_4': '弩弓手',
    'char_5': '武士',
    'char_6': '黑道士'
}


class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setGeometry(100, 100, 800, 600)
        self.setWindowTitle("Mir4 Ranking")
        self.setWindowIcon(QIcon('../icon/shark.ico'))
        layout = QVBoxLayout()
        self.setLayout(layout)

        # Create a group box for the server input elements
        group_box = QGroupBox("Server Selection")
        group_layout = QVBoxLayout()
        group_box.setLayout(group_layout)
        layout.addWidget(group_box)

        # Use a form layout to align the server label and input field
        form_layout = QFormLayout()
        group_layout.addLayout(form_layout)

        self.server_label = QLabel("Enter server names (separated by commas):", self)
        form_layout.addRow(self.server_label)

        self.server_input = QLineEdit(self)
        self.server_input.setText(
            "ASIA081,ASIA021")
        form_layout.addRow(self.server_input)

        # add radio let user choose to get character ranking or clan ranking
        self.character_ranking = QRadioButton("Character Ranking", self)
        self.character_ranking.setChecked(True)
        self.clan_ranking = QRadioButton("Clan Ranking", self)
        self.clan_ranking.setChecked(False)
        # two radio buttons 並排並靠左
        self.radio_layout = QHBoxLayout()
        self.radio_layout.addWidget(self.character_ranking)
        self.radio_layout.addWidget(self.clan_ranking)
        self.radio_layout.addStretch(1)
        group_layout.addLayout(self.radio_layout)

        self.textbox = QTextEdit(self)
        self.textbox.setReadOnly(True)  # 設置為只讀
        layout.addWidget(self.textbox)

        # Add some spacing between the widgets
        layout.addSpacing(20)

        # Set a fixed size for the message label
        self.message_label = QLabel("", self)
        self.message_label.setFixedHeight(20)
        layout.addWidget(self.message_label)

        self.get_rankings_button = QPushButton("Get Rankings", self)
        layout.addWidget(self.get_rankings_button)
        self.get_rankings_button.clicked.connect(self.get_rankings_button_clicked)
        self.load_settings()

    def load_settings(self):
        settings = QSettings("Summer Sharky", "Sharky's Mir4 Ranking")
        previous_server_input = settings.value("server_input")
        if previous_server_input is not None:
            self.server_input.setText(previous_server_input)

    def save_settings(self):
        settings = QSettings("Summer Sharky", "Sharky's Mir4 Ranking")
        settings.setValue("server_input", self.server_input.text())

    def closeEvent(self, event):
        self.save_settings()
        super().closeEvent(event)

    def print_message(self, message):
        # 將訊息逐行打印到QPlainTextEdit中
        self.textbox.append(message)

    def update_message_label(self, message):
        self.message_label.setText(message)
        QApplication.processEvents()

    def get_rankings_button_clicked(self):
        try:
            server_names = self.server_input.text().split(",")
            self.message_label.setText("Retrieving rankings...")
            self.textbox.clear()

            # Use threading to avoid freezing the GUI
            thread = threading.Thread(target=self.get_ranking_data_and_export, args=[server_names], daemon=True)
            thread.start()
        except Exception as e:
            self.print_message("Error: " + str(e))

    def get_ranking_data_and_export(self, target_server_list):
        self.get_rankings_button.setDisabled(True)
        error_list = self.get_several_ranking_data_then_export(target_server_list)
        if not error_list:
            self.message_label.setText("Rankings retrieved successfully.")
            self.print_message("Rankings retrieved successfully.")
        else:
            self.message_label.setText("Error retrieving rankings for the following servers: " + ", ".join(error_list))
            self.print_message("Error retrieving rankings for the following servers: " + ", ".join(error_list))
        self.get_rankings_button.setDisabled(False)

    def data_to_csv(self, data, filename):
        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                if self.character_ranking.isChecked():
                    fieldnames = ['Rank', 'Username', 'PowerScore', 'Class', 'Clan']
                else:
                    fieldnames = ['Rank', 'Clan', 'ClanLeader', 'ClanPowerScore']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

                # Write header row
                writer.writeheader()

                # Write data rows
                for row in data:
                    writer.writerow(row)
        except Exception as e:
            self.print_message(f"Error: {e}")

    def get_several_ranking_data_then_export(self, target_server_list: list):
        now = datetime.now()
        date_string = now.strftime("%Y%m%d")
        error_list = []
        servers = self.get_all_server_para()
        data_type = "character" if self.character_ranking.isChecked() else "clan"
        target_servers = {}
        for server_name in target_server_list:
            if server_name not in servers:
                self.print_message(f"Error: {server_name} does not exist.")
                error_list.append(server_name)
                continue
            target_servers[server_name] = servers[server_name]

        for server_name, server_para in target_servers.items():
            self.print_message(f"Getting data from {server_name}")
            try:
                csv_data = self.get_ranking_data(server_para['world_group_id'], server_para['world_id'], 1, 10)
            except AttributeError as e:
                self.print_message(f"Error: {server_name} Fail.")
                error_list.append(server_name)
                continue
            # 將結果存入csv檔案
            filename = f"{server_name}_{data_type}_Ranking_{date_string}.csv"

            self.data_to_csv(csv_data, filename)  # Write to csv file
            self.print_message(f"Success: {server_name} Done.")

        return error_list

    def get_ranking_data(self, world_group_id, world_id, start_page, end_page):
        for attempt in range(10):
            try:
                # radio button 1: character ranking
                # radio button 2: clan ranking
                if self.character_ranking.isChecked():
                    return self.fetch_character_ranking(end_page, start_page, world_group_id, world_id)
                elif self.clan_ranking.isChecked():
                    return self.fetch_clan_ranking(end_page, start_page, world_group_id, world_id)

            except Exception as e:
                if attempt < 9:
                    self.print_message(f"Failed, Retrying... ({attempt + 1}/10)")
                else:
                    raise e

    def fetch_clan_ranking(self, end_page, start_page, world_group_id, world_id):
        result_list = []
        for page in range(start_page, end_page + 1):
            url = f"https://forum.mir4global.com/rank?ranktype=4&worldgroupid={world_group_id}&worldid={world_id}&classtype=&searchname=&page={page}"
            response = requests.get(url)
            soup = BeautifulSoup(response.text, "html.parser")
            result = soup.find("tbody", {'id': 'lists'})

            # 取出想要的內容
            for row in result.find_all('tr'):
                rank = row.find('span', {'class': 'num'}).text.strip()
                clan = row.find_all('td', {'class': None})[0].find('span').text.strip()
                leader = row.find('span', {'class': 'user_name'}).text.strip()
                clan_points = row.find('td', {'class': 'text_right'}).text.strip()
                result_list.append(
                    {'Rank': rank,  'Clan': clan, 'ClanLeader': leader, 'ClanPowerScore': clan_points})
        return result_list

    def fetch_character_ranking(self, end_page, start_page, world_group_id, world_id):
        result_list = []
        for page in range(start_page, end_page + 1):
            url = f"https://forum.mir4global.com/rank?ranktype=1&worldgroupid={world_group_id}&worldid={world_id}&classtype=&searchname=&page={page}"
            response = requests.get(url)
            soup = BeautifulSoup(response.text, "html.parser")
            result = soup.find("tbody", {'id': 'lists'})

            # 取出想要的內容
            for row in result.find_all('tr'):
                rank = row.find('span', {'class': 'num'}).text.strip()
                username = row.find('span', {'class': 'user_name'}).text.strip()
                points = row.find('td', {'class': 'text_right'}).text.strip()
                clan = row.find_all('td', {'class': None})[1].find('span').text.strip()
                png_name = re.search(r'char_\d+', row.find('span', {'class': 'user_icon'})['style']).group(0)
                class_name = png_to_class.get(png_name, 'Unknown')  # Look up class name in dictionary
                result_list.append(
                    {'Rank': rank, 'Username': username, 'PowerScore': points, 'Class': class_name, 'Clan': clan})
        return result_list

    def get_all_server_para(self):
        response = requests.get(
            'https://forum.mir4global.com/rank?ranktype=1')
        soup = BeautifulSoup(response.text, "html.parser")
        data_dict = {}

        for ul in soup.find_all('div', {'class': 'depth2 world'}):
            for li in ul.find_all('li'):
                link = li.find('a')
                if link:
                    world_group_id = int(re.findall(r"'([^']*)'", link['href'])[0])
                    world_id = int(re.findall(r"'([^']*)'", link['href'])[2])
                    data_dict[link.text.strip()] = {'world_group_id': world_group_id, 'world_id': world_id}
        return data_dict


if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='dark_teal.xml')
    window = Window()
    window.show()
    sys.exit(app.exec_())
