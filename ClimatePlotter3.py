from geopy.geocoders import Nominatim
import pandas as pd
from mpl_toolkits.basemap import Basemap
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import cm
import sys
import os
import time
import shutil

matplotlib.use('QtAgg')
from matplotlib.patches import Polygon
# from matplotlib.collections import PatchCollection
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt6.QtWidgets import *
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import *


# from matplotlib.figure import Figure

class LectureMapApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # self.excelFilePath = './Plotter_Output/ClimatePlotter.xlsx'
        # self.plotPath = './Plotter_Output/'
        self.excelFilePath = os.path.join(os.path.dirname(__file__), 'Plotter_Output', 'ClimatePlotter.xlsx')
        self.viewsFilePath = os.path.join(os.path.dirname(__file__), 'Views.xlsx')
        self.plotPath = 'Plotter_Output'

        self.setWindowTitle("Lecture Map Plotter")
        self.setGeometry(100, 100, 1200, 900)

        self.df_views = self.read_views_file(self.viewsFilePath)

        self.stateList = [
            'Baden-Württemberg',
            'Bayern',
            'Berlin',
            'Brandenburg',
            'Bremen',
            'Hamburg',
            'Hessen',
            'Mecklenburg-Vorpommern',
            'Niedersachsen',
            'Nordrhein-Westfalen',
            'Rheinland-Pfalz',
            'Saarland',
            'Sachsen',
            'Sachsen-Anhalt',
            'Schleswig-Holstein',
            'Thüringen'
        ]

        self.initUI()

    def get_address(self, name):
        geolocator = Nominatim(user_agent="http")
        location = geolocator.geocode(name + ", Germany")
        if location:
            return location.address
        else:
            return None

    def get_coordinates(self, address, city_name, state_name, plz_code):
        geolocator = Nominatim(user_agent="http")
        location = geolocator.geocode(
            address + ", " + city_name + ", " + state_name + ", " + str(plz_code) + ", Germany")
        if location:
            return location.latitude, location.longitude
        else:
            return None, None

    def read_excel_file(self, file_path):
        try:
            xl = pd.ExcelFile(file_path)
            if len(xl.sheet_names) > 1 and 'Events' in xl.sheet_names and 'Stats' in xl.sheet_names:
                df_events = xl.parse('Events')
                df_stats = xl.parse('Stats')
            else:
                if 'Events' not in xl.sheet_names:
                    df_events = pd.DataFrame(
                        columns=['Datum', 'Schul/Uni Name', 'Adresse', 'Stadt', 'Bundesland', 'PLZ', 'Tische', 'Teilnehmer'])
                    df_stats = pd.DataFrame(columns=['Schul/Uni Name', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                                     'EventCount', 'CityEventTotal', 'TotalTables',
                                                     'TotalParticipants', 'CityParticipantsTotal'])
                    msgText = "Events sheet not found, Stats reset"
                elif 'Stats' not in xl.sheet_names:
                    df_events = xl.parse('Events')
                    df_stats = pd.DataFrame(columns=['Schul/Uni Name', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                                     'EventCount', 'CityEventTotal', 'TotalTables',
                                                     'TotalParticipants', 'CityParticipantsTotal'])
                    msgText = "Stats sheet not found, created new from values in Events sheet"
                    self.recalculateStatistics(df_events, df_stats)
                self.create_msg_box("Sheet Not Found", msgText, 'warning')
            return df_events, df_stats
        except FileNotFoundError:
            df_events = pd.DataFrame(
                columns=['Datum', 'Schul/Uni Name', 'Adresse', 'Stadt', 'Bundesland', 'PLZ', 'Tische', 'Teilnehmer'])
            df_stats = pd.DataFrame(columns=['Schul/Uni Name', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                             'EventCount', 'CityEventTotal', 'TotalTables',
                                             'TotalParticipants', 'CityParticipantsTotal'])
            with pd.ExcelWriter(self.excelFilePath) as writer:
                df_events.to_excel(writer, sheet_name='Events', index=False)
                df_stats.to_excel(writer, sheet_name='Stats', index=False)
                self.create_msg_box("File Not Found",
                                    "ClimatePlotter.xlsx not found in the Plotter_Output directory, created new file",
                                    'warning')
            return df_events, df_stats

    def update_excel(self, file_path, date, name, address, city, state, plz, lat, lon, tables=0,
                     participants=0):
        df_events, df_stats = self.read_excel_file(file_path)
        df_events.loc[len(df_events.index) + 1] = [date, name, address, city, state, str(plz), int(tables), int(participants)]
        # successFlag = self.recalculateStatistics(df_events, df_stats, lat, lon)
        # return successFlag
        try:
            with pd.ExcelWriter(self.excelFilePath) as writer:
                df_events.to_excel(writer, sheet_name='Events', index=False)
                df_stats.to_excel(writer, sheet_name='Stats', index=False)
            return True
        except OSError:
            return False

    def plot_map(self, df_events, df_stats, save_path, canvas, lat, lon, llc_lat, llc_lon, urc_lat, urc_lon,
                 view='Deutschland', doSave=False):
        canvas.figure.clf()
        ax = canvas.figure.add_subplot(111)
        '''
        resolution: c (crude), l (low), i (intermediate), h (high), f (full) or None
        projection: 'merc' (Mercator), 'cyl' (Cylindrical Equidistant), 'mill' (Miller Cylindrical), 'gall' (Gall Stereographic Cylindrical), 'cea' (Cylindrical Equal Area), 'lcc' (Lambert Conformal), 'tmerc' (Transverse Mercator), 'omerc' (Oblique Mercator), 'nplaea' (North-Polar Lambert Azimuthal), 'npaeqd' (North-Polar Azimuthal Equidistant), 'nplaea' (South-Polar Lambert Azimuthal), 'spaeqd' (South-Polar Azimuthal Equidistant), 'aea' (Albers Equal Area), 'stere' (Stereographic), 'robin' (Robinson), 'eck4' (Eckert IV), 'eck6' (Eckert VI), 'kav7' (Kavrayskiy VII), 'mbtfpq' (McBryde-Thomas Flat-Polar Quartic), 'sinu' (Sinusoidal), 'gall' (Gall Stereographic Cylindrical), 'hammer' (Hammer), 'moll' (Mollweid
        espg: 3857 (Web Mercator)
        '''
        m = Basemap(resolution='h', lat_0=(urc_lat - llc_lat)/2, lon_0=(urc_lon - llc_lon)/2, llcrnrlon=llc_lon, llcrnrlat=llc_lat,
                    urcrnrlon=urc_lon, urcrnrlat=urc_lat, epsg=3857)
        m.drawcountries()

        '''
        https://gdz.bkg.bund.de/index.php/default/wmts-topplusopen-wmts-topplus-open.html
        web
        web_grau
        web_scale
        web_scale_grau
        web_light
        web_light_grau
        '''
        wms_server = 'https://sgx.geodatenzentrum.de/wms_topplus_open?request=GetCapabilities&service=wms'
        m.wmsimage(wms_server, layers=["web_light"], verbose=False)

        m.drawcoastlines()
        if view == 'Deutschland' or view in self.stateList:
            '''
            Handle drawing Germany or the states
            '''
            shapePath = os.path.join(os.path.dirname(__file__), 'shapefiles', 'DEU_adm1')
            m.readshapefile(shapePath, 'areas')
            df_poly = pd.DataFrame(columns=['shapes', 'area'])
            for info, shape in zip(m.areas_info, m.areas):
                shape_array = np.array(shape)
                df_poly.loc[len(df_poly.index) + 1] = [Polygon(shape_array), info['NAME_1']]

            df1 = df_stats.groupby('CityParticipantsTotal')
            colors = iter(cm.winter(np.linspace(1, 0, len(df1.groups))))

            for group_cluster, data in df1:
                msize = group_cluster * 5
                if msize > 100:
                    msize = 100
                x, y = m(data['Longitude'], data['Latitude'])
                ax.scatter(x, y, label=group_cluster, marker='*', s=msize, color=next(colors))
        else:
            '''
            Handle drawing the cities
            '''
            df1 = df_stats.loc[df_stats['Stadt'] == view].reset_index(drop=True)
            df_max = max(df1['TotalParticipants'])
            for group_cluster, data in df1.iterrows():
                msize = data['TotalParticipants'] * 5
                if msize > 200:
                    msize = 200
                col = cm.winter(data['TotalParticipants'] / df_max)
                x, y = m(data['Longitude'], data['Latitude'])
                ax.scatter(x, y, label=group_cluster, marker='*', s=msize, color=col)
        canvas.draw()
        if doSave:
            save_path = os.path.join(os.path.dirname(__file__), save_path,
                                     f'{pd.Timestamp.now().strftime("%Y-%m-%d_H%HM%MS%S")}_{view}.png')
            plt.savefig(save_path, format='png', dpi=300)


    def drawInitialMap(self):
        df_events, df_stats = self.read_excel_file(self.excelFilePath)
        self.plot_map(df_events,
                      df_stats,
                      self.plotPath,
                      self.canvas,
                      self.df_views.loc[self.df_views['View'] == 'Deutschland']['lat_0'],
                      self.df_views.loc[self.df_views['View'] == 'Deutschland']['lon_0'],
                      self.df_views.loc[self.df_views['View'] == 'Deutschland']['llcrnrlat'],
                      self.df_views.loc[self.df_views['View'] == 'Deutschland']['llcrnrlon'],
                      self.df_views.loc[self.df_views['View'] == 'Deutschland']['urcrnrlat'],
                      self.df_views.loc[self.df_views['View'] == 'Deutschland']['urcrnrlon']
                      )

    def read_views_file(self, file_path):
        return pd.read_excel(file_path)

    def addExcelFiles(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        dialog.setNameFilter("Excel files (*.xlsx)")
        if dialog.exec():
            fileNames = dialog.selectedFiles()
            for fileName in fileNames:
                self.bulkImportList.addItem(fileName)

    def removeExcelFiles(self):
        try:
            self.bulkImportList.takeItem(self.bulkImportList.currentRow())
        except OSError:
            self.create_msg_box("File Error", "Error in removing files", 'warning')

    def clearExcelFiles(self):
        self.bulkImportList.clear()

    def bulkImportExcelFiles(self):
        for i in reversed(range(self.bulkImportList.count())):
            bulkFilePath = self.bulkImportList.item(i).text()
            df = pd.read_excel(bulkFilePath)
            df = df.fillna('')  # Replace NaN with empty string
            for index, row in df.iterrows():
                date = row['Datum'].strftime("%d.%m.%Y")
                name = row['Schul/Uni']
                address = str(row['Adresse'])
                city = str(row['Stadt'])
                state = str(row['Bundesland'])
                plzCode = str(row['PLZ'])
                table = str(row['Tische'])
                participant = str(row['Teilnehmer'])
                latitude, longitude = self.get_coordinates(address, city, state, plzCode)
                time.sleep(0.25)  # rate limit to max 4 requests per second
                self.update_excel(self.excelFilePath, date, name, address, city, state, plzCode, latitude, longitude,
                                  table, participant)
            if os.path.exists(bulkFilePath):
                os.remove(bulkFilePath)
            self.bulkImportList.takeItem(i)
        df_events, df_stats = self.read_excel_file(self.excelFilePath)
        self.recalculateStatistics(df_events, df_stats)
        self.drawInitialMap()

    def onLookupAddressButtonClicked(self):
        name = self.nameEdit.text()
        # create exception handler to handle address of NoneType
        try:
            raw_address = self.get_address(name).split(', ')
            self.nameEdit.setText(raw_address[0])
            self.plzEdit.setText(raw_address[-2])
            self.cityEdit.setText(raw_address[-4])
            state = raw_address[-3]
        except AttributeError:
            self.create_msg_box("Address Not Found", "Address not found. Please try again.", 'warning')
            return

        # find the state in the state list and set the combo box to that index
        self.stateCombo.setCurrentIndex(self.stateList.index(state))

        substring_list = ['straße',
                          'platz',
                          'weg',
                          'gasse',
                          'allee',
                          'ring',
                          'Straße',
                          'Platz',
                          'Weg',
                          'Gasse',
                          'Allee',
                          'Ring',
                          'Im ',
                          "Am ",
                          "An ",
                          "Auf ",
                          "In ",
                          "Zum ",
                          "Zur ",
                          "Zu ",
                          "Bei ",
                          "Unter ",
                          "Über ",
                          "Vor ",
                          "Hinter ",
                          "Neben ",
                          "stieg",
                          "Stieg",
                          "steg",
                          "Steg",
                          "steig",
                          "Steig",
                          "Ufer"
                          ]
        for field in raw_address:
            if any(substring in field for substring in substring_list):
                number = raw_address.index(field)
                if number == 2:
                    self.addressEdit.setText(raw_address[2] + ' ' + raw_address[1])
                else:
                    self.addressEdit.setText(raw_address[1])

    def clearAll(self):
        self.nameEdit.setText('')
        self.addressEdit.setText('')
        self.cityEdit.setText('')
        self.plzEdit.setText('')
        self.dateEdit.setDate(QDate.currentDate())
        self.stateCombo.setCurrentIndex(0)
        self.tableEdit.setValue(0)
        self.participantEdit.setValue(0)

    def onRecalculateButtonClicked(self):
        msgBox = QMessageBox(self)
        msgBox.setIcon(QMessageBox.Icon.Warning)
        msgBox.setWindowTitle("Recalculate Statistics?")
        msgBox.setText(
            "Proceeding will clear all existing statistics and recalculate them from scratch.\n\n "
            "Are you sure you want to proceed?\n\n "
            "Calculation may take a while.")
        msgBox.setStandardButtons(
            QMessageBox.StandardButton.Yes |
            QMessageBox.StandardButton.No
        )
        button = msgBox.exec()

        if button == QMessageBox.StandardButton.No:
            return
        elif button == QMessageBox.StandardButton.Yes:
            df_events, df_stats = self.read_excel_file(self.excelFilePath)
            df_stats = pd.DataFrame(columns=['Schul/Uni Name', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                             'EventCount', 'CityEventTotal', 'TotalTables',
                                             'TotalParticipants', 'CityParticipantsTotal'])
            self.recalculateStatistics(df_events, df_stats)
            self.create_msg_box("Complete", "Recalculation Complete")
            return

    def recalculateStatistics(self, df_events, df_stats, lat=None, lon=None):
        df_stats = pd.DataFrame(columns=['Schul/Uni Name', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                         'EventCount', 'CityEventTotal', 'TotalTables',
                                         'TotalParticipants', 'CityParticipantsTotal'])
        for index, row in df_events.iterrows():
            name = row['Schul/Uni Name']
            address = row['Adresse']
            city = row['Stadt']
            state = row['Bundesland']
            plz = row['PLZ']
            tables = row['Tische']
            participants = row['Teilnehmer']
            school_exists = ((df_stats['Schul/Uni Name'] == name) & (df_stats['Stadt'] == city)).any().any()
            if school_exists:
                df_stats.loc[(df_stats['Schul/Uni Name'] == name) & (df_stats['Stadt'] == city), 'EventCount'] += 1
                df_stats.loc[
                    (df_stats['Schul/Uni Name'] == name) & (df_stats['Stadt'] == city), 'TotalTables'] += int(
                    tables)
                df_stats.loc[
                    (df_stats['Schul/Uni Name'] == name) & (df_stats['Stadt'] == city), 'TotalParticipants'] += int(
                    participants)
            else:
                if lat is None or lon is None:
                    lat, lon = self.get_coordinates(address, city, state, plz)
                    time.sleep(0.25)  # rate limit to max 4 requests per second
                df_stats.loc[len(df_stats.index) + 1] = [name, city, plz, lat, lon, 1, -1, int(tables),
                                                         int(participants), -1]
                lat = None
                lon = None
            city_event_total = df_stats[df_stats['Stadt'] == city]['EventCount'].sum()
            df_stats.loc[df_stats['Stadt'] == city, 'CityEventTotal'] = int(city_event_total)
            city_participants_total = df_stats[df_stats['Stadt'] == city]['TotalParticipants'].sum()
            df_stats.loc[df_stats['Stadt'] == city, 'CityParticipantsTotal'] = int(city_participants_total)
        try:
            with pd.ExcelWriter(self.excelFilePath) as writer:
                df_events.to_excel(writer, sheet_name='Events', index=False)
                df_stats.to_excel(writer, sheet_name='Stats', index=False)
            return True
        except OSError:
            return False

    def onArchiveButtonClicked(self):

        msgBox = QMessageBox(self)
        msgBox.setIcon(QMessageBox.Icon.Warning)
        msgBox.setWindowTitle("Archive Excel?")
        msgBox.setText(
            "Archive and clear to start a new data file\n\n"
            "Archive and keep to keep the current data file\n\n"
            "Cancel to keep the current data file without archiving"
        )
        archive_and_clear_button = QPushButton("Archive and Clear")
        archive_and_keep_button = QPushButton("Archive and Keep")

        msgBox.addButton(archive_and_clear_button, QMessageBox.ButtonRole.AcceptRole)
        msgBox.addButton(archive_and_keep_button, QMessageBox.ButtonRole.AcceptRole)
        msgBox.addButton(QMessageBox.StandardButton.Cancel)

        msgBox.exec()

        if msgBox.clickedButton() == archive_and_clear_button:
            self.archive_and_clear()
        elif msgBox.clickedButton() == archive_and_keep_button:
            self.archive_and_keep()
        elif msgBox.clickedButton() == QMessageBox.StandardButton.Cancel:
            return


    def archive_and_clear(self):
        try:
            self.archiveExcel()
            df_events = pd.DataFrame(
                columns=['Datum', 'Schul/Uni Name', 'Adresse', 'Stadt', 'Bundesland', 'PLZ', 'Tische', 'Teilnehmer'])
            df_stats = pd.DataFrame(columns=['Schul/Uni Name', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                             'EventCount', 'CityEventTotal', 'TotalTables',
                                             'TotalParticipants', 'CityParticipantsTotal'])
            with pd.ExcelWriter(self.excelFilePath) as writer:
                df_events.to_excel(writer, sheet_name='Events', index=False)
                df_stats.to_excel(writer, sheet_name='Stats', index=False)
        except OSError:
            self.create_msg_box("Error", "Error in archiving data")
        return

    def archive_and_keep(self):
        self.archiveExcel()
        return

    def archiveExcel(self):
        date = pd.Timestamp.now().strftime("%Y-%m-%d_H%HM%MS%S")
        folder_path = os.path.join(os.path.dirname(__file__), 'Plotter_Output', 'Archive')
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        shutil.copy(
            self.excelFilePath,
            os.path.join(
                folder_path,
                date + "_Archived_ClimatePlotter.xlsx"
            )
        )
        self.create_msg_box("Complete", "Archiving Complete")
        return

    def initUI(self):

        self.setGeometry(300, 300, 1024, 700)
        self.setWindowTitle('Deutschland - Klimapuzzle Plotter')
        self.setWindowIcon(QIcon('icon/icon.png'))
        # Create central widget
        centralWidget = QWidget(self)
        self.setCentralWidget(centralWidget)

        # Layout
        mylayout = QHBoxLayout()
        leftBox = QGridLayout()
        middleBox = QVBoxLayout()
        rightBox = QGridLayout()

        '''
        BulkButtonBox
        '''
        self.addButton = QPushButton("Add", self)
        self.removeButton = QPushButton("Remove", self)
        self.clearButton = QPushButton("Clear", self)

        '''
        BulkImportBox
        '''
        self.bulkImportList = QListWidget(self)
        bulkImportGroupBox = QGroupBox("Bulk Import", self)
        BulkImportBoxLayout = QGridLayout()
        bulkImportGroupBox.setLayout(BulkImportBoxLayout)

        self.bulkImportButton = QPushButton("Bulk Import", self)

        BulkImportBoxLayout.addWidget(self.bulkImportList, 0, 0, 3, 1)
        BulkImportBoxLayout.addWidget(self.addButton, 0, 1)
        BulkImportBoxLayout.addWidget(self.removeButton, 1, 1)
        BulkImportBoxLayout.addWidget(self.clearButton, 2, 1)
        BulkImportBoxLayout.addWidget(self.bulkImportButton, 3, 0, 1, 2)

        self.addButton.clicked.connect(self.addExcelFiles)
        self.removeButton.clicked.connect(self.removeExcelFiles)
        self.clearButton.clicked.connect(self.clearExcelFiles)
        self.bulkImportButton.clicked.connect(self.bulkImportExcelFiles)

        '''
        Input Box
        '''
        # Add widgets to layout
        self.nameLabel = QLabel("School/University Name:", self)
        self.nameEdit = QLineEdit(self)

        self.addressLookupButton = QPushButton("Lookup Address", self)
        self.addressLookupButton.clicked.connect(self.onLookupAddressButtonClicked)

        self.addressLabel = QLabel("Address:", self)
        self.addressEdit = QLineEdit(self)

        self.cityLabel = QLabel("City:", self)
        self.cityEdit = QLineEdit(self)

        self.plzLabel = QLabel("PLZ Code (Optional):", self)
        self.plzEdit = QLineEdit(self)

        self.stateLabel = QLabel("State:", self)
        self.stateCombo = QComboBox()
        for state in self.stateList:
            self.stateCombo.addItem(state)

        self.dateLabel = QLabel("Date (dd.MM.yyyy):", self)
        self.dateEdit = QDateEdit(self)
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.setDisplayFormat("dd.MM.yyyy")
        self.dateEdit.setDate(QDate.currentDate())
        minDate = QDate(2020, 1, 1)
        maxDate = QDate(2099, 12, 31)
        self.dateEdit.setMinimumDate(minDate)
        self.dateEdit.setMaximumDate(maxDate)

        self.tableLabel = QLabel("Tische:", self)
        self.tableEdit = QSpinBox(self)
        self.tableEdit.setMinimum(0)

        self.participantLabel = QLabel("Teilnehmer:", self)
        self.participantEdit = QSpinBox(self)
        self.participantEdit.setMinimum(0)

        # New button for updating CSV
        self.updateCsvButton = QPushButton("Update Excel", self)
        self.updateCsvButton.clicked.connect(self.onUpdateCsvButtonClicked)

        self.plotButton = QPushButton("Update Plot", self)
        self.plotButton.clicked.connect(lambda: self.onPlotButtonClicked(doSave=True))

        self.clearAllButton = QPushButton("Clear All", self)
        self.clearAllButton.clicked.connect(self.clearAll)

        self.recalculateButton = QPushButton("Recalculate Statistics", self)
        self.recalculateButton.clicked.connect(self.onRecalculateButtonClicked)

        self.archiveButton = QPushButton("Archive Xlsx", self)
        self.archiveButton.clicked.connect(self.onArchiveButtonClicked)

        InputBoxGroupBox = QGroupBox("Manual Input", self)
        InputBoxLayout = QGridLayout()
        InputBoxGroupBox.setLayout(InputBoxLayout)

        InputBoxLayout.addWidget(self.nameLabel, 0, 0)
        InputBoxLayout.addWidget(self.nameEdit, 0, 1)
        InputBoxLayout.addWidget(self.addressLookupButton, 1, 0, 1, 2)
        InputBoxLayout.addWidget(self.addressLabel, 2, 0)
        InputBoxLayout.addWidget(self.addressEdit, 2, 1)
        InputBoxLayout.addWidget(self.cityLabel, 3, 0)
        InputBoxLayout.addWidget(self.cityEdit, 3, 1)
        InputBoxLayout.addWidget(self.plzLabel, 4, 0)
        InputBoxLayout.addWidget(self.plzEdit, 4, 1)
        InputBoxLayout.addWidget(self.stateLabel, 5, 0)
        InputBoxLayout.addWidget(self.stateCombo, 5, 1)
        InputBoxLayout.addWidget(self.dateLabel, 6, 0)
        InputBoxLayout.addWidget(self.dateEdit, 6, 1)
        InputBoxLayout.addWidget(self.tableLabel, 7, 0)
        InputBoxLayout.addWidget(self.tableEdit, 7, 1)
        InputBoxLayout.addWidget(self.participantLabel, 8, 0)
        InputBoxLayout.addWidget(self.participantEdit, 8, 1)
        InputBoxLayout.addWidget(self.updateCsvButton, 9, 0)
        InputBoxLayout.addWidget(self.plotButton, 9, 1)
        InputBoxLayout.addWidget(self.clearAllButton, 10, 0, 1, 2)

        FileControlGroupBox = QGroupBox("File Control", self)
        FileControlBoxLayout = QVBoxLayout()
        FileControlGroupBox.setLayout(FileControlBoxLayout)
        FileControlBoxLayout.addWidget(self.recalculateButton)
        FileControlBoxLayout.addWidget(self.archiveButton)

        '''
        middleBox
        '''
        self.figure = plt.figure()
        self.canvas = FigureCanvas(self.figure)
        middleBox.addWidget(self.canvas)  # Add the canvas to the layout

        '''
        rightBox
        '''
        self.plotListWidget = QListWidget(self)
        self.setupPlotListWidget()
        self.plotListWidget.clicked.connect(self.onplotListWidgetClicked)

        viewInputBox = QGridLayout()

        viewInputGroupBox = QGroupBox("Manual View Input", self)
        viewInputBoxLayout = QGridLayout()
        viewInputGroupBox.setLayout(viewInputBoxLayout)
        self.viewNameLabel = QLabel("View Name:", self)
        self.viewName = QLineEdit(self)
        self.viewLatLabel = QLabel("Central Latitude:", self)
        self.viewLat = QLineEdit(self)
        self.viewLonLabel = QLabel("Central Longitude:", self)
        self.viewLon = QLineEdit(self)
        self.viewLLCLatLabel = QLabel("Lower Left Corner Latitude:", self)
        self.viewLLCLat = QLineEdit(self)
        self.viewLLCLonLabel = QLabel("Lower Left Corner Longitude:", self)
        self.viewLLCLon = QLineEdit(self)
        self.viewURCLatLabel = QLabel("Upper Right Corner Latitude:", self)
        self.viewURCLat = QLineEdit(self)
        self.viewURCLonLabel = QLabel("Upper Right Corner Longitude:", self)
        self.viewURCLon = QLineEdit(self)
        self.cityLookupButton = QPushButton("Find City Coordinates", self)

        viewInputBoxLayout.addWidget(self.viewNameLabel, 0, 0)
        viewInputBoxLayout.addWidget(self.viewName, 0, 1)
        viewInputBoxLayout.addWidget(self.viewLatLabel, 1, 0)
        viewInputBoxLayout.addWidget(self.viewLat, 1, 1)
        viewInputBoxLayout.addWidget(self.viewLonLabel, 2, 0)
        viewInputBoxLayout.addWidget(self.viewLon, 2, 1)
        viewInputBoxLayout.addWidget(self.viewLLCLatLabel, 3, 0)
        viewInputBoxLayout.addWidget(self.viewLLCLat, 3, 1)
        viewInputBoxLayout.addWidget(self.viewLLCLonLabel, 4, 0)
        viewInputBoxLayout.addWidget(self.viewLLCLon, 4, 1)
        viewInputBoxLayout.addWidget(self.viewURCLatLabel, 5, 0)
        viewInputBoxLayout.addWidget(self.viewURCLat, 5, 1)
        viewInputBoxLayout.addWidget(self.viewURCLonLabel, 6, 0)
        viewInputBoxLayout.addWidget(self.viewURCLon, 6, 1)

        cityLookupGroupBox = QGroupBox("City Lookup", self)
        cityLookupBoxLayout = QGridLayout()
        cityLookupGroupBox.setLayout(cityLookupBoxLayout)

        self.cityNameLabel = QLabel("City Name:", self)
        self.cityName = QLineEdit(self)
        self.stateNameLabel = QLabel("State Name:", self)
        self.stateLookupCombo = QComboBox()
        for state in self.stateList:
            self.stateLookupCombo.addItem(state)
        self.latitudeHeightLabel = QLabel("Latitude Height:", self)
        self.latitudeHeight = QLineEdit(self)
        self.latitudeHeight.setText("0.3")
        self.longitudeWidthLabel = QLabel("Longitude Width:", self)
        self.longitudeWidth = QLineEdit(self)
        self.longitudeWidth.setText("0.3")
        self.cityLookupButton.clicked.connect(self.onCityLookupButtonClicked)
        cityLookupBoxLayout.addWidget(self.cityNameLabel, 0, 0)
        cityLookupBoxLayout.addWidget(self.cityName, 0, 1)
        cityLookupBoxLayout.addWidget(self.stateNameLabel, 1, 0)
        cityLookupBoxLayout.addWidget(self.stateLookupCombo, 1, 1)
        cityLookupBoxLayout.addWidget(self.latitudeHeightLabel, 2, 0)
        cityLookupBoxLayout.addWidget(self.latitudeHeight, 2, 1)
        cityLookupBoxLayout.addWidget(self.longitudeWidthLabel, 3, 0)
        cityLookupBoxLayout.addWidget(self.longitudeWidth, 3, 1)
        cityLookupBoxLayout.addWidget(self.cityLookupButton, 4, 0, 1, 2)

        self.PreViewAddButton = QPushButton("Preview View", self)
        self.PreViewAddButton.clicked.connect(self.onPreViewButtonClicked)
        self.ViewAddButton = QPushButton("Save View", self)
        self.ViewAddButton.clicked.connect(self.onViewAddButtonClicked)
        self.ViewRemoveButton = QPushButton("Remove View", self)
        self.ViewRemoveButton.clicked.connect(self.onViewRemoveButtonClicked)

        viewInputBox.addWidget(viewInputGroupBox, 0, 0, 5, 2)
        viewInputBox.addWidget(cityLookupGroupBox, 5, 0, 5, 2)
        viewInputBox.addWidget(self.PreViewAddButton, 10, 0, 1, 2)
        viewInputBox.addWidget(self.ViewAddButton, 11, 0, 1, 2)
        viewInputBox.addWidget(self.ViewRemoveButton, 12, 0, 1, 2)

        rightBox.addWidget(self.plotListWidget, 0, 0, 5, 1)
        rightBox.addLayout(viewInputBox, 6, 0, 3, 1)

        leftBox.addWidget(bulkImportGroupBox, 0, 0, 3, 2)
        leftBox.addWidget(InputBoxGroupBox, 3, 0, 3, 2)
        leftBox.addWidget(FileControlGroupBox, 6, 0, 1, 2)

        mylayout.addLayout(leftBox)
        mylayout.addLayout(middleBox)
        mylayout.addLayout(rightBox)

        centralWidget.setLayout(mylayout)

        # Draw the initial map
        self.drawInitialMap()

    def create_msg_box(self, title, text, type='info'):
        msgBox = QMessageBox(self)
        if type == 'info':
            type = QMessageBox.Icon.Information
        elif type == 'warning':
            type = QMessageBox.Icon.Warning
        elif type == 'critical':
            type = QMessageBox.Icon.Critical
        msgBox.setIcon(QMessageBox.Icon.Information)
        msgBox.setWindowTitle(title)
        msgBox.setText(text)
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.exec()

    def onUpdateCsvButtonClicked(self):
        date = self.dateEdit.text()
        name = self.nameEdit.text()
        address = self.addressEdit.text()
        city = self.cityEdit.text()
        plzCode = self.plzEdit.text()
        state = self.stateCombo.currentText()
        tables = self.tableEdit.text()
        participants = self.participantEdit.text()
        if name == '' or city == '' or state == '':
            self.create_msg_box('Input Empty!', 'Fill in all the fields!\n\nData not saved.', 'warning')
        else:
            latitude, longitude = self.get_coordinates(address, city, state, plzCode)
            successFlag = self.update_excel(
                self.excelFilePath,
                date,
                name,
                address,
                city,
                state,
                plzCode,
                latitude,
                longitude,
                tables,
                participants
            )
            df_events, df_stats = self.read_excel_file(self.excelFilePath)
            self.recalculateStatistics(df_events, df_stats)
            if not successFlag:
                self.create_msg_box('Input Failed!',
                                    'Data not saved.\n\nPlease check that the Excel file not open and try again.',
                                    'warning')
            else:
                self.create_msg_box('Input Successful!',
                                    'Data saved successfully!\n\nClick on "Update Plot" to see the changes.')
                # Clear inputs
                self.clearAll()


    def getPlotItem(self):
        entry = self.df_views.loc[self.df_views['View'] == self.plotListWidget.currentItem().text()]
        if not entry.empty:
            entry = entry.iloc[0]  # Select the first row
        return entry

    def onPlotButtonClicked(self, doSave=False):
        df_events, df_stats = self.read_excel_file(self.excelFilePath)
        plotItem = self.getPlotItem()
        self.plot_map(
            df_events,
            df_stats,
            self.plotPath,
            self.canvas,
            plotItem['lat_0'],
            plotItem['lon_0'],
            plotItem['llcrnrlat'],
            plotItem['llcrnrlon'],
            plotItem['urcrnrlat'],
            plotItem['urcrnrlon'],
            plotItem['View'],
            doSave
        )

    def setupPlotListWidget(self):
        self.plotListWidget.clear()
        self.plotListWidget.addItems(self.df_views['View'].tolist())
        self.plotListWidget.setCurrentRow(0)

    def onplotListWidgetClicked(self):
        entry = self.getPlotItem()
        self.viewName.setText(entry['View'])
        self.viewLat.setText(str(entry['lat_0']))
        self.viewLon.setText(str(entry['lon_0']))
        self.viewLLCLat.setText(str(entry['llcrnrlat']))
        self.viewLLCLon.setText(str(entry['llcrnrlon']))
        self.viewURCLat.setText(str(entry['urcrnrlat']))
        self.viewURCLon.setText(str(entry['urcrnrlon']))

    def ClearViewText(self):
        self.viewName.clear()
        self.viewLat.clear()
        self.viewLon.clear()
        self.viewLLCLat.clear()
        self.viewLLCLon.clear()
        self.viewURCLat.clear()
        self.viewURCLon.clear()

    def onPreViewButtonClicked(self):
        lat = self.viewLat.text()
        lon = self.viewLon.text()
        llc_lat = self.viewLLCLat.text()
        llc_lon = self.viewLLCLon.text()
        urc_lat = self.viewURCLat.text()
        urc_lon = self.viewURCLon.text()
        view_name = self.viewName.text()
        if lat == '' or lon == '' or llc_lat == '' or llc_lon == '' or urc_lat == '' or urc_lon == '' or view_name == '':
            self.create_msg_box('Input Empty!', 'Fill in all the fields!\n\nData not saved.', 'warning')
        else:
            df_events, df_stats = self.read_excel_file(self.excelFilePath)
            self.plot_map(
                df_events,
                df_stats,
                self.plotPath,
                self.canvas,
                float(lat),
                float(lon),
                float(llc_lat),
                float(llc_lon),
                float(urc_lat),
                float(urc_lon),
                self.viewName.text(),
                doSave=False
            )

    def onCityLookupButtonClicked(self):
        lat, lon = self.get_coordinates('', self.cityName.text(), self.stateLookupCombo.currentText(), '')
        self.viewName.setText(self.cityName.text())
        self.viewLat.setText(str(lat))
        self.viewLon.setText(str(lon))
        self.viewLLCLat.setText(str(lat - float(self.latitudeHeight.text())/2))
        self.viewLLCLon.setText(str(lon - float(self.longitudeWidth.text())/2))
        self.viewURCLat.setText(str(lat + float(self.latitudeHeight.text())/2))
        self.viewURCLon.setText(str(lon + float(self.longitudeWidth.text())/2))
        self.onPreViewButtonClicked()

    def onViewAddButtonClicked(self):
        lat = self.viewLat.text()
        lon = self.viewLon.text()
        llc_lat = self.viewLLCLat.text()
        llc_lon = self.viewLLCLon.text()
        urc_lat = self.viewURCLat.text()
        urc_lon = self.viewURCLon.text()
        view_name = self.viewName.text()
        if lat == '' or lon == '' or llc_lat == '' or llc_lon == '' or urc_lat == '' or urc_lon == '' or view_name == '':
            self.create_msg_box('Input Empty!', 'Fill in all the fields!\n\nData not saved.', 'warning')
        else:
            view_exists = ((self.df_views['View'] == view_name)).any().any()
            if view_exists:
                self.df_views.loc[self.df_views['View'] == view_name, 'lat_0'] = float(lat)
                self.df_views.loc[self.df_views['View'] == view_name, 'lon_0'] = float(lon)
                self.df_views.loc[self.df_views['View'] == view_name, 'llcrnrlat'] = float(llc_lat)
                self.df_views.loc[self.df_views['View'] == view_name, 'llcrnrlon'] = float(llc_lon)
                self.df_views.loc[self.df_views['View'] == view_name, 'urcrnrlat'] = float(urc_lat)
                self.df_views.loc[self.df_views['View'] == view_name, 'urcrnrlon'] = float(urc_lon)
                self.df_views.to_excel(self.viewsFilePath, index=False)
                self.setupPlotListWidget()
                self.ClearViewText()
                self.create_msg_box('Input Successful!',
                                    'View updated successfully')
            else:
                self.df_views.loc[len(self.df_views.index) + 1] = [view_name, float(lat), float(lon), float(llc_lon), float(llc_lat), float(urc_lon), float(urc_lat)]
                self.df_views.to_excel(self.viewsFilePath, index=False)
                self.setupPlotListWidget()
                self.ClearViewText()
                self.create_msg_box('Input Successful!',
                                    'View saved successfully')

    def onViewRemoveButtonClicked(self):
        entry = self.plotListWidget.currentItem().text()
        if entry == 'Deutschland' or entry in self.stateList:
            self.create_msg_box('Error',
                                'Cannot remove default views',
                                'warning')
            return
        else:
            self.df_views = self.df_views[self.df_views.View != entry]
            self.df_views.to_excel(self.viewsFilePath, index=False)
            self.setupPlotListWidget()
            self.ClearViewText()
            self.create_msg_box('Success',
                                'View removed successfully')


def run_app():
    app = QApplication([])
    mainWin = LectureMapApp()
    mainWin.show()
    # app.exec()
    sys.exit(app.exec())


if __name__ == '__main__':
    run_app()
