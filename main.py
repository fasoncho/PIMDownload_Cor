import os
import requests
import pandas as pd
import openpyxl
import re
import math
import urllib

from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *

import sys
from bs4 import BeautifulSoup

country_websites = pd.read_excel('website_data.xlsx')
country_websites_dict = country_websites.set_index('Country').to_dict()['Url']



class ProgressBar(QProgressBar):

    def __init__(self, title):
        QProgressBar.__init__(self)
        layout = QVBoxLayout()
        self.setLayout(layout)


class CountrySelector(QWidget):

    def __init__(self, title, country_website_base):
        QWidget.__init__(self)
        layout = QHBoxLayout()
        self.setLayout(layout)
        self.downloadURL = '/ww/en/'

        self.label = QLabel()
        self.label.setText(title)
        self.label.setFixedWidth(130)
        self.label.setFont(QFont("Arial"))
        self.label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(self.label)

        self.countryList = QComboBox()
        self.countryList.addItems(country_websites_dict.keys())

        self.countryList.currentIndexChanged.connect(self.setCountry)
        layout.addWidget(self.countryList)
        layout.addStretch()

    def setCountry(self, i):
        self.downloadURL = list(country_websites_dict.values())[i]

    def setLabelWidth(self, width):
        self.label.setFixedWidth(width)
        # --------------------------------------------------------------------

    def setlineEditWidth(self, width):
        self.lineEdit.setFixedWidth(width)


class FileBrowser(QWidget):
    OpenFile = 0
    OpenFiles = 1
    OpenDirectory = 2
    SaveFile = 3

    def __init__(self, title, mode=OpenFile):
        QWidget.__init__(self)
        layout = QHBoxLayout()
        self.setLayout(layout)
        self.enable = False
        self.browser_mode = mode
        self.filter_name = ' Excel files (*.xlsx; *.xls)'
        self.dirpath = QDir.currentPath()

        self.label = QLabel()
        self.label.setText(title)
        self.label.setFixedWidth(130)
        self.label.setFont(QFont("Arial"))
        self.label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(self.label)

        self.lineEdit = QLineEdit(self)
        self.lineEdit.setFixedWidth(180)
        layout.addWidget(self.lineEdit)

        self.button = QPushButton('Browse')
        self.button.clicked.connect(self.getFile)
        layout.addWidget(self.button)
        layout.addStretch()

        self.filepaths = []

    def setMode(mode):
        self.mode = mode

    # --------------------------------------------------------------------
    # For example,
    #    setFileFilter('Images (*.png *.xpm *.jpg)')
    def setFileFilter(text):
        self.filter_name = text
        # --------------------------------------------------------------------

    def setDefaultDir(path):
        self.dirpath = path

    # --------------------------------------------------------------------
    def getFile(self):
        if self.browser_mode == FileBrowser.OpenFile:
            self.filepaths.append(QFileDialog.getOpenFileName(self, caption='Choose File',
                                                              directory=self.dirpath,
                                                              filter=self.filter_name)[0])
        elif self.browser_mode == FileBrowser.OpenDirectory:
            self.filepaths.append(QFileDialog.getExistingDirectory(self, caption='Choose Directory',
                                                                   directory=self.dirpath))
        else:
            return

        if len(self.filepaths) == 0:
            return
        else:
            self.lineEdit.setText(self.filepaths[0])
            return True
            # --------------------------------------------------------------------

    def setLabelWidth(self, width):
        self.label.setFixedWidth(width)
        # --------------------------------------------------------------------

    def setlineEditWidth(self, width):
        self.lineEdit.setFixedWidth(width)

    # --------------------------------------------------------------------
    def getPaths(self):
        return self.filepaths

    def on_text_changed(self):
        self.button.setEnabled(bool(self.lineEdit.text()))


# -------------------------------------------------------------------


class Demo(QDialog):
    def __init__(self, parent=None):
        QDialog.__init__(self, parent)

        # Ensure our window stays in front and give it a title
        self.secondary_endings = ['']
        self.proxies = None
        self.proxy_check = None
        self.button = None
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle("PIM Image Download")
        self.setFixedSize(440, 300)

        # Create and assign the main (vertical) layout.
        vlayout = QVBoxLayout()
        hlayout = QHBoxLayout()
        self.setLayout(vlayout)
        self.countryListPanel(vlayout)
        self.fileBrowserPanel(vlayout)
        self.radio = QRadioButton("PNG 4000 dpi", self)
        self.radio.toggled.connect(self.pngOrJpg)
        hlayout.addWidget(self.radio)
        # vlayout.addWidget(self.radio)
        self.file_format = 'jpg'
        self.link_ending = '_1500_jpg'
        self.radio = QRadioButton("JPG 1500 dpi", self)
        self.radio.setChecked(True)
        self.radio.toggled.connect(self.pngOrJpg)
        # vlayout.addWidget(self.radio)
        hlayout.addWidget(self.radio)
        vlayout.addStretch()
        self.addButtonPanel(hlayout)
        self.addExitButtonPanel(hlayout)
        vlayout.addLayout(hlayout)
        self.progressBar(vlayout)

        self.show()

    # --------------------------------------------------------------------
    def progressBar(self, parent_layout):

        vlayout = QVBoxLayout()
        self.progressBarWidget = ProgressBar('Select country')
        vlayout.addWidget(self.progressBarWidget)
        self.progressBarWidget.setGeometry(50, 100, 250, 30)
        self.progressBarWidget.setValue(0)
        vlayout.addStretch()
        parent_layout.addLayout(vlayout)

    def countryListPanel(self, parent_layout):

        vlayout = QHBoxLayout()
        self.countryListWidget = CountrySelector('Select country', country_websites_dict)
        vlayout.addWidget(self.countryListWidget)
        vlayout.addStretch()
        parent_layout.addLayout(vlayout)

    def fileBrowserPanel(self, parent_layout):
        vlayout = QVBoxLayout()

        self.fileFB = FileBrowser('Select Reference List', FileBrowser.OpenFile)
        self.dirFB = FileBrowser('Save in folder', FileBrowser.OpenDirectory)

        vlayout.addWidget(self.fileFB)
        vlayout.addWidget(self.dirFB)

        vlayout.addStretch()
        parent_layout.addLayout(vlayout)

    # --------------------------------------------------------------------
    def addButtonPanel(self, parentLayout):
        hlayout = QHBoxLayout()
        hlayout.addStretch()
        self.button = QPushButton("OK")
        self.button.clicked.connect(self.proxyCheck)
        self.button.clicked.connect(self.buttonAction)
        hlayout.addWidget(self.button)
        parentLayout.addLayout(hlayout)

    def addExitButtonPanel(self, parentLayout):
        hlayout = QHBoxLayout()
        hlayout.addStretch()
        self.button = QPushButton("Exit")
        self.button.clicked.connect(self.buttonExit)
        hlayout.addWidget(self.button)
        parentLayout.addLayout(hlayout)

    def pngOrJpg(self):

        if self.sender().text() == 'PNG 4000 dpi':
            self.link_ending = '_4000_png'
            self.file_format = 'png'
        else:
            self.link_ending = '_1500_jpg'
            self.file_format = 'jpg'

    # --------------------------------------------------------------------
    def proxyCheck(self):

        try:
            requests.get('https://eref.se.com/ww/en/product/A9F84416', proxies=self.proxies)
            self.proxy_check = True
        except OSError:
            self.proxy_check = False
            pass
        if self.proxy_check:
            self.proxies = {'http': 'http://gateway.schneider.zscaler.net:80',
                            'https': 'http://gateway.schneider.zscaler.net:80'}
        else:
            self.proxies = {}

    def buttonAction(self):

        value = 0
        self.refs_path = self.fileFB.getPaths()
        try:
            self.refs_path = os.path.abspath(self.refs_path[0])

        except IndexError:
            dialog = QMessageBox(parent=self, text="Please select reference file!")
            dialog.setWindowTitle("System message")
            ret = dialog.exec()
            return

        self.destination_path = self.dirFB.getPaths()
        try:
            self.destination_path = os.path.abspath(self.destination_path[0])
        except IndexError:
            dialog = QMessageBox(parent=self, text="Please select destination folder!")
            dialog.setWindowTitle("System message")
            ret = dialog.exec()
            return
        try:
            ref_data = pd.read_excel(self.refs_path, header=None)
        except ValueError:
            print('Select Excel file with references in first column')
            return
        not_found_counter = 0
        self.proxies = None
        for ref in ref_data.values:
            value += math.ceil(100 / len(ref_data))
            link = f'https://eref.se.com{self.countryListWidget.downloadURL}product/{ref[0]}'

            try:
                a = requests.get(link, proxies=self.proxies)
                soup_data = BeautifulSoup(requests.get(link, proxies=self.proxies).content, 'html.parser')
            except OSError:
                not_found_counter += 1
                continue

            try:
                doc_url_soup = soup_data.find('img', {'class': 'thumbnail'})
                doc_url = doc_url_soup['src']
            except TypeError:
                not_found_counter += 1
                continue
            headers = {
                'User-Agent': 'PIM_Image_Downloader 1.0',
            }
            doc_id = re.search(r'(Doc_Ref=)(.+)(&p)', doc_url)[2]
            img_download_url = f'https://download.se.com/files?p_Doc_Ref={doc_id}&p_File_Type=rendition{self.link_ending}&default_image=DefaultProductImage.png'

            try:
                img_file = requests.get(img_download_url, proxies=self.proxies, headers=headers).content
            except OSError:
                not_found_counter += 1
                return

            with open(f"{self.destination_path}\\{ref[0]}.{self.file_format}", 'wb') as handler:
                handler.write(img_file)

            self.progressBarWidget.setValue(int(value))

        if not_found_counter != 0:
            dialog = QMessageBox(parent=self, text=f'{not_found_counter} references not found')
            dialog.setWindowTitle("System message")
            ret = dialog.exec()
        else:
            dialog = QMessageBox(parent=self, text='All product images were successfully downloaded')
            dialog.setWindowTitle("System message")
            ret = dialog.exec()

    def buttonExit(self):
        sys.exit(0)


# ========================================================
if __name__ == '__main__':
    # Create the Qt Application
    app = QApplication(sys.argv)
    demo = Demo()  # <<-- Create an instance
    demo.show()
    sys.exit(app.exec())
