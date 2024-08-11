from geopy.geocoders import *
import pandas as pd
from mpl_toolkits.basemap import Basemap
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import cm
import sys
import os
import shutil

matplotlib.use('QtAgg')
from matplotlib.patches import Polygon
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PySide6.QtWidgets import *
from PySide6.QtGui import QIcon
from PySide6.QtCore import *


class AddressError(Exception):
    '''
    Exception class
    '''
    pass


class LectureMapApp(QMainWindow):
    """
    A GUI application for visualizing and managing lecture and event data on a map.

    The `LectureMapApp` class provides a user interface built with Pyside6 that allows users to interact with
    event and view data stored in Excel files. Users can add, edit, remove, and plot lecture or event locations
    on a map. The application supports custom views, allows for the import of event data, and provides the ability
    to save and manage these views for future use.

    Attributes:
        excelFilePath (str): The file path to the main ClimatePlotter Excel document where event data is stored.
        viewsFilePath (str): The file path to the Views Excel document where map view configurations are stored.
        plotPath (str): The directory path where plot output files are saved.
        df_views (pd.DataFrame): A DataFrame containing view data loaded from the Views Excel file.
        stateList (list of str): A list of German states used for state selection and validation within the application.
        canvas (QWidget): The canvas where map plots are rendered.
        plotListWidget (QListWidget): A widget displaying the list of available views for selection.
        nameEdit (QLineEdit): A text field for entering the name of the event or institution.
        addressEdit (QLineEdit): A text field for entering the address of the event or institution.
        cityEdit (QLineEdit): A text field for entering the city of the event or institution.
        plzEdit (QLineEdit): A text field for entering the postal code (PLZ) of the event or institution.
        stateCombo (QComboBox): A dropdown for selecting the state of the event or institution.
        tableEdit (QLineEdit): A text field for entering the number of tables at the event.
        participantEdit (QLineEdit): A text field for entering the number of participants at the event.
        viewLat (QLineEdit): A text field for entering the latitude of the map view.
        viewLon (QLineEdit): A text field for entering the longitude of the map view.
        viewLLCLat (QLineEdit): A text field for entering the lower-left corner latitude of the map view.
        viewLLCLon (QLineEdit): A text field for entering the lower-left corner longitude of the map view.
        viewURCLat (QLineEdit): A text field for entering the upper-right corner latitude of the map view.
        viewURCLon (QLineEdit): A text field for entering the upper-right corner longitude of the map view.
        viewName (QLineEdit): A text field for entering the name of the map view.
        df_events (pd.DataFrame): A DataFrame containing event data loaded from the ClimatePlotter Excel file.

    Methods:
        __init__(): Initializes the application, sets up file paths, loads initial data, and configures the user interface.
        initUI(): Sets up the user interface components and layouts for the application.
        read_views_file(file_path): Reads the views from the specified Excel file and returns them as a DataFrame.
        read_excel_file(file_path): Reads event data from the specified Excel file and returns it as a DataFrame.
        get_coordinates(city, state, plz, address=None): Retrieves the latitude and longitude coordinates for a given location.
        plot_map(df_events, df_stats, plotPath, canvas, lat, lon, llc_lat, llc_lon, urc_lat, urc_lon, view_name, doSave=False): Plots the event data on a map.
        onLookupAddressButtonClicked(): Handles the event when the 'Lookup Address' button is clicked.
        onPreViewButtonClicked(): Handles the event when the 'Pre-View' button is clicked.
        onViewAddButtonClicked(): Handles the event when the 'Add View' button is clicked.
        onCityLookupButtonClicked(): Handles the event when the 'City Lookup' button is clicked.
        onViewRemoveButtonClicked(): Handles the event when the 'Remove View' button is clicked.
        recalculateStatistics(df_events, df_stats, lat=None, lon=None): Recalculates statistics for the events and updates the Excel file.
        getPlotItem(): Retrieves the plot item from the DataFrame based on the currently selected item in the plot list widget.
        create_msg_box(title, message, icon='information'): Creates and displays a message box with the specified title, message, and icon.
        clearAll(): Clears all input fields in the user interface.
        setupPlotListWidget(): Sets up and populates the plot list widget with available views.

    """
    def __init__(self):
        super().__init__()

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
        '''
        Uses OSM's Nominatim to get a German address for a particular search name.
        country_codes set to de to limit search to Germany.
        addressdetails set to True to get more information about the address OSM has for the search object

        Args:
            name (str): Search string, ex: "HKA" or "Karlsruhe Institute of Technology".

        Returns:
            dict: geopy.location.Location.raw - dictionary containing unparsed location information returned
                from Nominatim.
        '''
        geolocator = Nominatim(user_agent="http")
        location = geolocator.geocode(name, country_codes="de", addressdetails=True)
        if location is not None:
            return location.raw
        else:
            self.create_msg_box("Address Error",
                                f"Error: {'Bad search results returned'} \n\n"
                                f"{name} returned no results.",
                                'warning')
            return None

    def get_coordinates(self, city_name, state_name, plz_code, address=''):
        '''
        Uses OSM's Nominatim to get latitude and longitude coordinates for the search object

        Args:
           city_name (str): City to be searched. Ex: "Regensburg"
           state_name (str): State to be searched. Ex: "Bayern"
           plz_code (str): Post code. Ex: "76139"
           address (str): Street Address of search object. Defaults to ''.

        Returns:
            str, str: latitude and longitude from Nominatim's location dict.
        '''
        geolocator = Nominatim(user_agent="http")
        location = geolocator.geocode(
            address + ", " + city_name + ", " + state_name + ", " + str(plz_code) + ", Germany")
        if location:
            return location.latitude, location.longitude
        else:
            return None, None

    def read_excel_file(self, file_path):
        '''
        Reads the ClimatePlotter excel file, parses Events and Stats, or if they do not exist, creates them.
        If the ClimatePlotter excel document does not exist, creates one.

        Args:
            file_path (str): Where to find the excel sheet

        Returns:
            pd.Dataframe, pd.Dataframe: Events dataframe, Stats dataframe
        '''
        try:
            xl = pd.ExcelFile(file_path)
            if len(xl.sheet_names) > 1 and 'Events' in xl.sheet_names and 'Stats' in xl.sheet_names:
                df_events = xl.parse('Events')
                df_stats = xl.parse('Stats')
            else:
                if 'Events' not in xl.sheet_names:
                    df_events = pd.DataFrame(
                        columns=['Datum', 'Hochschule', 'Adresse', 'Stadt', 'Bundesland', 'PLZ', 'Tische',
                                 'Teilnehmer'])
                    df_stats = pd.DataFrame(columns=['Hochschule', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                                     'EventCount', 'CityEventTotal', 'TotalTables',
                                                     'TotalParticipants', 'CityParticipantsTotal'])
                    msgText = "Events sheet not found, Stats reset"
                elif 'Stats' not in xl.sheet_names:
                    df_events = xl.parse('Events')
                    df_stats = pd.DataFrame(columns=['Hochschule', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                                     'EventCount', 'CityEventTotal', 'TotalTables',
                                                     'TotalParticipants', 'CityParticipantsTotal'])
                    msgText = "Stats sheet not found, created new from values in Events sheet"
                    self.recalculateStatistics(df_events, df_stats)
                self.create_msg_box("Sheet Not Found", msgText, 'warning')
            return df_events, df_stats
        except FileNotFoundError:
            df_events = pd.DataFrame(
                columns=['Datum', 'Hochschule', 'Adresse', 'Stadt', 'Bundesland', 'PLZ', 'Tische', 'Teilnehmer'])
            df_stats = pd.DataFrame(columns=['Hochschule', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
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
        """
        Writes the new rows of data to the ClimatePlotter Excel document.

        Args:
            file_path (str): The path to the ClimatePlotter document.
            date (str): The date of the event being added.
            name (str): The name of the school or university.
            address (str): The street address of the school or university.
            city (str): The city of the school or university.
            state (str): The state of the school or university.
            plz (str): The postal code of the school or university.
            lat (str): [UNUSED] The latitude coordinates of the school or university.
            lon (str): [UNUSED] The longitude coordinates of the school or university.
            tables (int): The number of tables at the event.
            participants (int): The total number of participants at the event.

        Returns:
            bool: A flag indicating whether the write to the Excel document was successful.
        """
        df_events, df_stats = self.read_excel_file(file_path)
        df_events.loc[len(df_events.index) + 1] = [date, name, address, city, state, str(plz), int(tables),
                                                   int(participants)]
        try:
            with pd.ExcelWriter(self.excelFilePath) as writer:
                df_events.to_excel(writer, sheet_name='Events', index=False)
                df_stats.to_excel(writer, sheet_name='Stats', index=False)
            return True
        except OSError:
            return False

    def plot_map(self, df_events, df_stats, save_path, canvas, lat, lon, llc_lat, llc_lon, urc_lat, urc_lon,
                 view='Deutschland', doSave=False):
        '''
        Map plotting tool using Basemap.

        Args:
            df_events (pd.Dataframe): [UNUSED] Dataframe of events
            df_stats (pd.Dataframe): Dataframe of stats relating to how many events a school/city have held.
            save_path (str): Location to save the generated map image.
            canvas (FigureCanvasQTAgg object): Matplotlib backend object as a plot canvas for drawing the map
            lat (float): [UNUSED] Latitude coordinate of plot center.
            lon (float): [UNUSED] Longitude coordinate of plot center
            llc_lat (float): Latitude - Lower left corner of view window
            llc_lon (float): Longitude - Lower left corner of view window
            urc_lat (float): Latitude - Upper right corner of view window
            urc_lon (float): Longitude - Upper right corner of view window
            view (str): Name of the view, used for title of saved map image. Defaults to 'Deutschland'
            doSave (bool): Triggers the saving of the plot image. Defaults to False

        Returns:
            None
        '''
        canvas.figure.clf()
        ax = canvas.figure.add_subplot(111)
        '''
        resolution: c (crude), l (low), i (intermediate), h (high), f (full) or None
        projection: 'merc' (Mercator), 'cyl' (Cylindrical Equidistant), 'mill' (Miller Cylindrical), 'gall' (Gall Stereographic Cylindrical), 'cea' (Cylindrical Equal Area), 'lcc' (Lambert Conformal), 'tmerc' (Transverse Mercator), 'omerc' (Oblique Mercator), 'nplaea' (North-Polar Lambert Azimuthal), 'npaeqd' (North-Polar Azimuthal Equidistant), 'nplaea' (South-Polar Lambert Azimuthal), 'spaeqd' (South-Polar Azimuthal Equidistant), 'aea' (Albers Equal Area), 'stere' (Stereographic), 'robin' (Robinson), 'eck4' (Eckert IV), 'eck6' (Eckert VI), 'kav7' (Kavrayskiy VII), 'mbtfpq' (McBryde-Thomas Flat-Polar Quartic), 'sinu' (Sinusoidal), 'gall' (Gall Stereographic Cylindrical), 'hammer' (Hammer), 'moll' (Mollweid
        espg: 3857 (Web Mercator)
        '''
        m = Basemap(resolution='h', lat_0=(urc_lat - llc_lat) / 2, lon_0=(urc_lon - llc_lon) / 2, llcrnrlon=llc_lon,
                    llcrnrlat=llc_lat,
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
                ax.scatter(x, y, label=group_cluster, marker='o', s=msize, color=next(colors))
        else:
            '''
            Handle drawing the cities
            '''
            df1 = df_stats.loc[df_stats['Stadt'] == view].reset_index(drop=True)
            if df1.empty:
                pass
            else:
                df_max = max(df1['TotalParticipants'])
                for group_cluster, data in df1.iterrows():
                    msize = data['TotalParticipants'] * 5
                    if msize > 200:
                        msize = 200
                    col = cm.winter(data['TotalParticipants'] / df_max)
                    x, y = m(data['Longitude'], data['Latitude'])
                    ax.scatter(x, y, label=group_cluster, marker='o', s=msize, color=col)
        canvas.draw()
        if doSave:
            save_path = os.path.join(os.path.dirname(__file__), save_path,
                                     f'{view}.png')
            plt.savefig(save_path, format='png', dpi=300)

    def drawInitialMap(self):
        '''
        Function to draw the initial map during UI initialization. Reads the ClimatePlotter excel document
        to add any existing events.

        Args:
            None

        Returns:
            None

        '''
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
        '''
        Function to read the existing view window coordinates from the views excel file.

        Args:
            file_path (str): Path to the Views excel file, usually found in the root directory.
        Returns:
            None

        '''
        return pd.read_excel(file_path)

    def addExcelFiles(self):
        '''
        Opens a dialog box, allowing the user to select input excel documents containing new events to be
        included in the running statistics. Files are added to the bulkImportList, which shows up in the
        list widget.

        Args:
            None

        Returns:
            None

        '''
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        dialog.setNameFilter("Excel files (*.xlsx)")
        if dialog.exec():
            fileNames = dialog.selectedFiles()
            for fileName in fileNames:
                self.bulkImportList.addItem(fileName)

    def removeExcelFiles(self):
        '''
        Allows the user to remove files from the bulkImportList, such as if they were added in error.

        Args:
            None

        Returns:
            None

        '''
        try:
            self.bulkImportList.takeItem(self.bulkImportList.currentRow())
        except OSError:
            self.create_msg_box("File Error", "Error in removing files", 'warning')

    def clearExcelFiles(self):
        '''
        Clears the bulkImportList completely.

        Args:
            None

        Returns:
            None

        '''
        self.bulkImportList.clear()

    def acceptBulkImport(self):
        '''
        Accept the import. Used in the import popup showing the file data vs the OSM data.

        Args:
            None

        Returns:
            None

        '''
        self.dialogBox.accept()
        self.dialogBox.close()

    def rejectBulkImport(self):
        '''
        Reject the import. Used in the import popup showing the file data vs the OSM data.

        Args:
            None

        Returns:
            None

        '''
        self.dialogBox.reject()
        self.dialogBox.close()

    def validate_fields(self):
        '''
        Validation of data fields. If all fields are filled, enable the button associated with
        acceptBulkImport

        Args:
            None

        Returns:
            None

        '''
        if (self.dateEdit.text() and
                self.nameEdit.text() and
                self.addressEdit.text() and
                self.cityEdit.text() and
                self.plzEdit.text() and
                self.tableEdit.value() > 0 and
                self.participantEdit.value() > 0):
            self.buttonBox.button(QDialogButtonBox.StandardButton.Ok).setEnabled(True)
        else:
            self.buttonBox.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)

    def bulkImportExcelFiles(self):
        '''
        Iterates through the bulkImportList/list widget, reads in the rows of each excel input document.
        Validates the data for completeness. Appends the data to the Events and Stats dataframes, calculates
        statistics and plots the results.

        Args:
            None

        Returns:
            None

        '''
        importedData = []

        for i in reversed(range(self.bulkImportList.count())):
            bulkFilePath = self.bulkImportList.item(i).text()

            # Load the Excel file into a DataFrame
            df = pd.read_excel(bulkFilePath)

            # Check for empty fields in any row
            if df.isnull().any(axis=1).any():
                self.create_msg_box(
                    "Validation Error",
                    f"File '{bulkFilePath.split('/')[-1]}' contains rows with empty fields.\nThis file will be skipped.",
                    'warning'
                )
                importedData.append((bulkFilePath, i, "Empty Fields"))
                continue

            # Fill NaN values with empty strings to avoid issues later
            df = df.fillna('')

            for index, row in df.iterrows():
                if not self.validate_essential_fields(row, bulkFilePath, index):
                    importedData.append((bulkFilePath, i, index))
                    continue

                plz = self.get_valid_plz(row, bulkFilePath, index)
                if plz is None:
                    importedData.append((bulkFilePath, i, index))
                    continue

                tables, participants = self.get_valid_table_participants(row, bulkFilePath, index)
                if tables is None or participants is None:
                    importedData.append((bulkFilePath, i, index))
                    continue

                raw_address = self.get_address_from_row(row, plz, bulkFilePath, index)
                if raw_address is None:
                    importedData.append((bulkFilePath, i, index))
                    continue

                self.process_and_display_dialog(row, raw_address, bulkFilePath, index, tables, participants)

        self.cleanup_imported_files(importedData)

        df_events, df_stats = self.read_excel_file(self.excelFilePath)
        df_events = df_events.fillna('')  # Replace NaN with empty string
        df_stats = df_stats.fillna('')  # Replace NaN with empty string
        self.recalculateStatistics(df_events, df_stats)
        self.drawInitialMap()

    def validate_essential_fields(self, row, bulkFilePath, index):
        '''
        Validates the data from the current row of the input excel document to ensure that key data is available.
        Without this data, Nominatim's search will return None, making it impossible to plot.

        Args:
            row (dict): Current row from the excel document
            bulkFilePath (str): Path to the current excel document
            index (int): Row index of the current row

        Returns:
            bool: Returns false if fields are missing, else returns true

        '''

        required_fields = ['Hochschule', 'Stadt', 'Bundesland']
        missing_fields = [field for field in required_fields if not row[field]]

        if missing_fields:
            self.create_msg_box(
                "Validation Error",
                f"Missing essential information in file: {bulkFilePath.split('/')[-1]}, Line: {index + 1}. "
                f"Required fields: {', '.join(missing_fields)}.",
                'warning'
            )
            return False
        return True

    def get_valid_plz(self, row, bulkFilePath, index):
        '''
        Validates the post code from the current row of the input excel document.

        Args:
           row (dict): Current row from the excel document
           bulkFilePath (str): Path to the current excel document
           index (int): Row index of the current row

        Returns:
            str or None: If the post code exists, returns the post code as a string. If the post code is invalid,
                returns None. If the post code does not exist, returns empty string ''.

        '''
        if row['PLZ']:
            try:
                return str(int(row['PLZ']))
            except ValueError:
                self.create_msg_box(
                    "Validation Error",
                    f"Invalid PLZ value in file: {bulkFilePath.split('/')[-1]}, Line: {index + 1}.",
                    'warning'
                )
                return None
        return ''

    def get_valid_table_participants(self, row, bulkFilePath, index):
        '''
        Validates the number of tables and participants.

        Args:
           row (dict): Current row from the excel document
           bulkFilePath (str): Path to the current excel document
           index (int): Row index of the current row

        Returns:
            int, int or None, None: If tables and participants are positive and non-zero, returns the
                integer values. If there's a value error, returns None for each.

        '''
        try:
            tables = int(row['Tische'])
            participants = int(row['Teilnehmer'])
            if tables < 0 or participants < 0:
                raise ValueError("Tables or Participants cannot be negative.")
            return tables, participants
        except (ValueError, KeyError):
            self.create_msg_box(
                "Validation Error",
                f"Invalid table or participant values in file: {bulkFilePath.split('/')[-1]}, Line: {index + 1}.",
                'warning'
            )
            return None, None

    def get_address_from_row(self, row, plz, bulkFilePath, index):
        '''
        Gets the address from the row using the name, city, state, and post code of the school/uni.

        Args:
           row (dict): Current row from the excel document
           plz: (str): Post code of the current row
           bulkFilePath (str): Path to the current excel document
           index (int): Row index of the current row

        Returns:
            dict or None: If the raw address is found by the get_address function, returns the dict object.
                If the address is not found by the function, or if there is a validation error, returns None.
        '''
        raw_address = self.get_address(row['Hochschule'] +
                                       ", " +
                                       row['Stadt'] +
                                       ", " +
                                       row['Bundesland'] +
                                       ", " +
                                       plz)

        if not raw_address or 'address' not in raw_address:
            self.create_msg_box(
                "Validation Error",
                f"Incomplete address data returned from OpenStreetMap for file: {bulkFilePath.split('/')[-1]}, Line: {index + 1}.",
                'warning'
            )
            return None

        if not all(key in raw_address['address'] for key in ('postcode', 'road')):
            self.create_msg_box(
                "Validation Error",
                f"Incomplete address data returned from OpenStreetMap for file: {bulkFilePath.split('/')[-1]}, Line: {index + 1}.",
                'warning'
            )
            return None

        return raw_address

    def process_and_display_dialog(self, row, raw_address, bulkFilePath, index, tables, participants):
        '''
        To assist the user in validating input data from the excel documents, this function displays the raw address
        found by Nominatim, side by side with the data from the excel document.

        Args:
           row (dict): Current row from the excel document
           raw_address (dict): Raw address returned by Nominatim's search
           bulkFilePath (str): Path to the current excel document
           index (int): Row index of the current row
           tables (int): Number of tables
           participants (int): Number of participants

        Returns:
            None

        '''
        # Prepare dialog fields and populate them with the row data and suggestions
        suggested_name = raw_address['name']
        suggested_plz = raw_address['address']['postcode']
        suggested_address = raw_address['address']['road']
        if 'house_number' in raw_address['address']:
            suggested_address += ' ' + raw_address['address']['house_number']
        suggested_city = raw_address['address'].get('city', '')
        suggested_state = raw_address['address'].get('state', '')

        # Prepare and show the dialog
        self.prepare_dialog(
            row,
            suggested_name,
            suggested_address,
            suggested_city,
            suggested_plz,
            suggested_state,
            bulkFilePath,
            index
        )

        button = self.dialogBox.exec()

        if button == 1:  # Accepted
            self.save_imported_data(row, raw_address, tables, participants, bulkFilePath, index)

    def prepare_dialog(self, row, suggested_name, suggested_address, suggested_city, suggested_plz, suggested_state,
                       bulkFilePath, index):
        '''
        Creates the popup UI that allows the user to inspect and compare the information from the input excel documents
        versus the data found by Nominatim's search

           row (dict): Current row from the excel document
           suggested_name (str): Name of the school/university, as found by Nominatim
           suggested_address (str): Address of the school/university, as found by Nominatim
           suggested_city (str): City of the school/university, as found by Nominatim
           suggested_plz (str): Post code of the school/university, as found by Nominatim
           suggested_state (str): State of the school/university, as found by Nominatim
           bulkFilePath (str): Path to the current excel document
           index: (int) Row index of the current row

        Returns:
            None

        '''
        # User certification dialog setup
        self.dialogBox = QDialog(self)
        self.dialogBox.setWindowTitle(f"Bulk Import - Certify Data {bulkFilePath.split('/')[-1]}, Line: {index + 1}")
        self.dialogBox.resize(400, 200)
        self.dialogBox.layout = QGridLayout()

        '''
        Define dialog box columns
        '''
        firstColumn = QLabel('')
        secondColumn = QLabel('Data from Excel')
        thirdColumn = QLabel('Data from OpenStreetMap')

        '''
        Define Labels for column 1 and Edits for column 2.
        
        Column 3 is suggested values and are read-only
        '''
        dateLabel = QLabel("Date:", self)
        self.dateEdit = QLineEdit(self)
        if row['Datum'] == '':
            self.dateEdit.setText(QDate.currentDate().toString("dd.MM.yyyy"))
            self.dateEdit.setStyleSheet('background-color: red')
        else:
            self.dateEdit.setText(row['Datum'].strftime("%d.%m.%Y"))
        dateSuggested = QLabel('')

        nameLabel = QLabel("Name:", self)
        self.nameEdit = QLineEdit(self)
        self.nameEdit.setText(row['Hochschule'].strip())
        nameSuggested = QLineEdit(suggested_name)
        nameSuggested.setReadOnly(True)

        addressLabel = QLabel("Address:", self)
        self.addressEdit = QLineEdit(self)
        self.addressEdit.setText(str(row['Adresse']).strip())
        addressSuggested = QLineEdit(suggested_address)
        addressSuggested.setReadOnly(True)

        cityLabel = QLabel("City:", self)
        self.cityEdit = QLineEdit(self)
        self.cityEdit.setText(row['Stadt'].strip())
        citySuggested = QLineEdit(suggested_city)
        citySuggested.setReadOnly(True)

        plzLabel = QLabel("PLZ Code:", self)
        self.plzEdit = QLineEdit(self)
        self.plzEdit.setText(str(row['PLZ']).strip())
        plzSuggested = QLineEdit(suggested_plz)
        plzSuggested.setReadOnly(True)

        stateLabel = QLabel("State:", self)
        stateCombo = QComboBox()
        for state in self.stateList:
            stateCombo.addItem(state)
        state = str(row['Bundesland']).strip()
        if state in self.stateList:
            stateCombo.setCurrentIndex(self.stateList.index(state))
        stateSuggested = QLineEdit(suggested_state)
        stateSuggested.setReadOnly(True)

        tableLabel = QLabel("Tables:", self)
        self.tableEdit = QSpinBox(self)
        self.tableEdit.setMinimum(0)
        self.tableEdit.setMaximum(2000)
        self.tableEdit.setValue(int(row['Tische']))
        tableSuggested = QLabel('')

        participantLabel = QLabel("Participants:", self)
        self.participantEdit = QSpinBox(self)
        self.participantEdit.setMinimum(0)
        self.participantEdit.setMaximum(9999)
        self.participantEdit.setValue(int(row['Teilnehmer']))
        participantSuggested = QLabel('')

        # Setup dialog and buttons
        QBtn = QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        self.buttonBox = QDialogButtonBox(QBtn)
        self.buttonBox.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)

        '''
        Define the widget layout for the data review dialog box
        '''
        self.dialogBox.layout.addWidget(firstColumn, 0, 0)
        self.dialogBox.layout.addWidget(secondColumn, 0, 1)
        self.dialogBox.layout.addWidget(thirdColumn, 0, 2)
        self.dialogBox.layout.addWidget(dateLabel, 1, 0)
        self.dialogBox.layout.addWidget(self.dateEdit, 1, 1)
        self.dialogBox.layout.addWidget(dateSuggested, 1, 2)
        self.dialogBox.layout.addWidget(nameLabel, 2, 0)
        self.dialogBox.layout.addWidget(self.nameEdit, 2, 1)
        self.dialogBox.layout.addWidget(nameSuggested, 2, 2)
        self.dialogBox.layout.addWidget(addressLabel, 3, 0)
        self.dialogBox.layout.addWidget(self.addressEdit, 3, 1)
        self.dialogBox.layout.addWidget(addressSuggested, 3, 2)
        self.dialogBox.layout.addWidget(cityLabel, 4, 0)
        self.dialogBox.layout.addWidget(self.cityEdit, 4, 1)
        self.dialogBox.layout.addWidget(citySuggested, 4, 2)
        self.dialogBox.layout.addWidget(plzLabel, 5, 0)
        self.dialogBox.layout.addWidget(self.plzEdit, 5, 1)
        self.dialogBox.layout.addWidget(plzSuggested, 5, 2)
        self.dialogBox.layout.addWidget(stateLabel, 6, 0)
        self.dialogBox.layout.addWidget(stateCombo, 6, 1)
        self.dialogBox.layout.addWidget(stateSuggested, 6, 2)
        self.dialogBox.layout.addWidget(tableLabel, 7, 0)
        self.dialogBox.layout.addWidget(self.tableEdit, 7, 1)
        self.dialogBox.layout.addWidget(tableSuggested, 7, 2)
        self.dialogBox.layout.addWidget(participantLabel, 8, 0)
        self.dialogBox.layout.addWidget(self.participantEdit, 8, 1)
        self.dialogBox.layout.addWidget(participantSuggested, 8, 2)
        self.dialogBox.layout.addWidget(self.buttonBox, 9, 0, 1, 3)
        self.dialogBox.setLayout(self.dialogBox.layout)

        # Connect signals to validation function
        self.dateEdit.textChanged.connect(self.validate_fields)
        self.nameEdit.textChanged.connect(self.validate_fields)
        self.addressEdit.textChanged.connect(self.validate_fields)
        self.cityEdit.textChanged.connect(self.validate_fields)
        self.plzEdit.textChanged.connect(self.validate_fields)
        self.tableEdit.valueChanged.connect(self.validate_fields)
        self.participantEdit.valueChanged.connect(self.validate_fields)
        self.validate_fields()

        self.buttonBox.accepted.connect(self.acceptBulkImport)
        self.buttonBox.rejected.connect(self.rejectBulkImport)

    def save_imported_data(self, row, raw_address, tables, participants, bulkFilePath, index):
        '''
        Pulls the input data from the text edit fields and saves it by calling update_excel.

        Args:
           row (dict): [UNUSED] Current row from the excel document
           raw_address: [UNUSED] (dict) Raw address returned by Nominatim's search
           tables (int): Number of tables
           participants (int): Number of Participants
           bulkFilePath (str): [UNUSED] Path to the current excel document
           index (int): [UNUSED] Row index of the current row

        Returns:
            None

        '''
        # Save the data
        date = self.dateEdit.text()
        name = self.nameEdit.text()
        address = self.addressEdit.text()
        city = self.cityEdit.text()
        state = self.stateCombo.currentText()
        plzCode = str(self.plzEdit.text())
        latitude, longitude = self.get_coordinates(city, state, plzCode, address)
        self.update_excel(self.excelFilePath, date, name, address, city, state, plzCode, latitude, longitude, tables,
                          participants)

    def cleanup_imported_files(self, importedData):
        '''
        If the file was imported properly then the excel document is deleted from the disk. If the file had errors,
        such as blank fields, it was not recorded, skipped, and is not deleted so that the user can correct the error.

        Args:
            importedData (list): List of files that were skipped during importing

        Returns:
            None

        '''
        for i in reversed(range(self.bulkImportList.count())):
            if i in [x[1] for x in importedData]:
                continue
            else:
                bulkFilePath = self.bulkImportList.item(i).text()
                if os.path.exists(bulkFilePath):
                    os.remove(bulkFilePath)
                self.bulkImportList.takeItem(i)

        if importedData:
            text = 'Not Processed:\n'
            for i in importedData:
                text += f'{i[0].split("/")[-1]} Line {i[1]}\n'
            self.create_msg_box("Bulk Import Complete", text, 'warning')

    def get_from_address(self, raw_address, *keys):
        """
        Helper function to retrieve the value from raw_address['address']
        using the first matching key in the provided keys.

        Args:
           raw_address (dict): Raw address returned by Nominatim's search
           keys (str): Keys to be checked if they exist in the address. Matching is necessary due to Nominatim's
                inconsistencies. Sometimes uses city, town, or village.
        Returns:
            str or None: If the key exists, return the value, otherwise return None
        """
        for key in keys:
            if key in raw_address['address']:
                return raw_address['address'][key]
        return None

    def onLookupAddressButtonClicked(self):
        """
        Handles the event when the 'Lookup Address' button is clicked.

        This method retrieves and processes the address information based on the name input by the user.
        It attempts to fetch and display the corresponding city, state, postcode, and road details in the appropriate UI fields.
        The method handles special cases for certain cities (Berlin, Hamburg, Bremen) and ensures that essential address components
        like the city, state, postcode, and road are available before populating the UI.

        Args:
            None

        Returns:
            None

        Raises:
            AddressError: If the city, state, postcode, or road information is missing from the retrieved address.

        """
        name = self.nameEdit.text()
        try:
            raw_address = self.get_address(name)
            if raw_address == None:
                return
            self.nameEdit.setText(raw_address['name'])

            # Attempt to retrieve the city, town, or village
            city = self.get_from_address(raw_address, 'city', 'town', 'village')
            if not city:
                raise AddressError("City/Town/Village not found in the address.")

            # Handle special cases for Berlin, Hamburg, and Bremen
            if city in ['Berlin', 'Hamburg', 'Bremen']:
                state = city
                self.cityEdit.setText(city)
            else:
                # Attempt to retrieve the state or equivalent
                state = self.get_from_address(raw_address, 'state', 'region')
                if not state:
                    raise AddressError("State/Region not found in the address.")
                self.cityEdit.setText(city)

            # Set the state in the combo box
            self.stateCombo.setCurrentIndex(self.stateList.index(state))

            # Handle postcode and road
            postcode = raw_address['address'].get('postcode')
            road = raw_address['address'].get('road')
            if postcode and road:
                self.plzEdit.setText(postcode)
                house_number = raw_address['address'].get('house_number', '')
                self.addressEdit.setText(f"{road} {house_number}".strip())
            else:
                raise AddressError("Incomplete address: postcode or road missing.")

        except AttributeError:
            self.create_msg_box("Address Not Found", "Address not found. Please try again.", 'warning')
            return
        except AddressError as e:
            self.create_msg_box("Address Error",
                                f"Error: {str(e)} \n\n"
                                f"Returned address was:\n\n {raw_address['display_name']}.",
                                'warning')
            return

    def clearAll(self):
        '''
        Clears all text input fields

        Args:
            None

        Returns:
            None

        '''
        self.nameEdit.setText('')
        self.addressEdit.setText('')
        self.cityEdit.setText('')
        self.plzEdit.setText('')
        self.dateEdit.setDate(QDate.currentDate())
        self.stateCombo.setCurrentIndex(0)
        self.tableEdit.setValue(0)
        self.participantEdit.setValue(0)

    def onRecalculateButtonClicked(self):
        '''
        When the button is clicked, recalculate the statistics for the Stats dataframe.

        Args:
            None

        Returns:
            None

        '''
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
            df_stats = pd.DataFrame(columns=['Hochschule', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                             'EventCount', 'CityEventTotal', 'TotalTables',
                                             'TotalParticipants', 'CityParticipantsTotal'])
            self.recalculateStatistics(df_events, df_stats)
            self.create_msg_box("Complete", "Recalculation Complete")
            return

    def recalculateStatistics(self, df_events, df_stats, lat=None, lon=None):
        """
        Recalculates statistical data based on event information and updates the statistics DataFrame.

        This method processes the events in the provided `df_events` DataFrame to generate and update statistical
        information in the `df_stats` DataFrame. The statistics include the number of events, total tables,
        and total participants for each university in the specified city. If latitude and longitude coordinates
        are not provided, the method attempts to retrieve them using the city's address information.

        Args:
            df_events (pd.DataFrame): A DataFrame containing event details, including fields such as 'Hochschule',
                'Adresse', 'Stadt', 'Bundesland', 'PLZ', 'Tische', and 'Teilnehmer'.
            df_stats (pd.DataFrame): A DataFrame that will be populated with statistics including 'Hochschule',
                'Stadt', 'PLZ', 'Latitude', 'Longitude', 'EventCount', 'CityEventTotal', 'TotalTables',
                'TotalParticipants', and 'CityParticipantsTotal'.
            lat (float, optional): Latitude coordinate to be used if not calculated or provided. Defaults to None.
            lon (float, optional): Longitude coordinate to be used if not calculated or provided. Defaults to None.

        Returns:
            bool: True if the Excel file was successfully written, otherwise False if an OSError occurs during the file writing process.

        Raises:
            None: All exceptions are handled internally, specifically file writing errors are caught and result in returning False.

        """
        df_stats = pd.DataFrame(columns=['Hochschule', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                         'EventCount', 'CityEventTotal', 'TotalTables',
                                         'TotalParticipants', 'CityParticipantsTotal'])
        for index, row in df_events.iterrows():
            name = row['Hochschule']
            address = row['Adresse']
            city = row['Stadt']
            state = row['Bundesland']
            plz = row['PLZ']
            tables = row['Tische']
            participants = row['Teilnehmer']
            school_exists = ((df_stats['Hochschule'] == name) & (df_stats['Stadt'] == city)).any().any()
            if school_exists:
                df_stats.loc[(df_stats['Hochschule'] == name) & (df_stats['Stadt'] == city), 'EventCount'] += 1
                df_stats.loc[
                    (df_stats['Hochschule'] == name) & (df_stats['Stadt'] == city), 'TotalTables'] += int(
                    tables)
                df_stats.loc[
                    (df_stats['Hochschule'] == name) & (df_stats['Stadt'] == city), 'TotalParticipants'] += int(
                    participants)
            else:
                if lat is None or lon is None:
                    lat, lon = self.get_coordinates(city, state, plz, address)
                    if lat is None or lon is None:
                        self.create_msg_box("Coordinates not found",
                                            f"No coordinates were found for {name} at {address}. {city}, {state}, {plz} \n"
                                            f"using coordinates for city instead.",
                                            'warning')
                        lat, lon = self.get_coordinates(city, state, plz)
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
        """
        Archives the existing ClimatePlotter excel document, copying or moving it to the archive folder.

        Args:
            None

        Returns:
            None

        """
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
        """
        Called from the onArchiveButtonClicked function. Moves the existing ClimatePlotter document to the
        archive folder and creates a new file.

        Args:
            None

        Returns:
            None

        """
        try:
            self.archiveExcel()
            df_events = pd.DataFrame(
                columns=['Datum', 'Hochschule', 'Adresse', 'Stadt', 'Bundesland', 'PLZ', 'Tische', 'Teilnehmer'])
            df_stats = pd.DataFrame(columns=['Hochschule', 'Stadt', 'PLZ', 'Latitude', 'Longitude',
                                             'EventCount', 'CityEventTotal', 'TotalTables',
                                             'TotalParticipants', 'CityParticipantsTotal'])
            with pd.ExcelWriter(self.excelFilePath) as writer:
                df_events.to_excel(writer, sheet_name='Events', index=False)
                df_stats.to_excel(writer, sheet_name='Stats', index=False)
        except OSError:
            self.create_msg_box("Error", "Error in archiving data")
        return

    def archive_and_keep(self):
        """
        Called from the onArchiveButtonClicked function. Copies the existing ClimatePlotter document to the
        archive folder.

        Args:
            None

        Returns:
            None

        """
        self.archiveExcel()
        return

    def archiveExcel(self):
        """
        Creates the copy of the ClimatePlotter excel document in the archive folder.

        Args:
            None

        Returns:
            None

        """
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
        """
        Initialize the UI of the main window

        Args:
            None

        Returns:
            None

        """
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

        self.plzLabel = QLabel("PLZ Code:", self)
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
        maxDate = QDate.currentDate()
        self.dateEdit.setMinimumDate(minDate)
        self.dateEdit.setMaximumDate(maxDate)

        self.tableLabel = QLabel("Tische:", self)
        self.tableEdit = QSpinBox(self)
        self.tableEdit.setMinimum(0)
        self.tableEdit.setMaximum(2000)
        self.tableEdit.setMinimum(0)

        self.participantLabel = QLabel("Teilnehmer:", self)
        self.participantEdit = QSpinBox(self)
        self.participantEdit.setMinimum(0)
        self.participantEdit.setMaximum(9999)

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
        self.cityLookupButton.clicked.connect(self.onCityLookupButtonClicked)
        cityLookupBoxLayout.addWidget(self.cityNameLabel, 0, 0)
        cityLookupBoxLayout.addWidget(self.cityName, 0, 1)
        cityLookupBoxLayout.addWidget(self.stateNameLabel, 1, 0)
        cityLookupBoxLayout.addWidget(self.stateLookupCombo, 1, 1)
        cityLookupBoxLayout.addWidget(self.cityLookupButton, 2, 0, 1, 2)

        self.PreViewAddButton = QPushButton("Preview View", self)
        self.PreViewAddButton.clicked.connect(self.onPreViewButtonClicked)
        self.ViewAddButton = QPushButton("Save View", self)
        self.ViewAddButton.clicked.connect(self.onViewAddButtonClicked)
        self.ViewRemoveButton = QPushButton("Remove View", self)
        self.ViewRemoveButton.clicked.connect(self.onViewRemoveButtonClicked)

        viewInputBox.addWidget(cityLookupGroupBox, 0, 0, 5, 2)
        viewInputBox.addWidget(viewInputGroupBox, 5, 0, 5, 2)
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
        """
        Creates popup style message boxes.

        Args:
            title (str): Title text for header for the pop up window
            text (str): Body text for the pop up window
            type (str): Class of pop up. Accepts 'info', 'warning', or 'critical'

        Returns:
            None

        """
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
        """
        Handles the event when the 'Update CSV' button is clicked.

        This method validates the input fields, converts relevant fields to integers,
        and checks that they are greater than zero. It then retrieves the coordinates
        based on the provided address details and updates the corresponding Excel file
        with the new data. If the data is successfully saved, the statistics are recalculated
        and the user is notified. If the input validation fails, or the Excel file cannot
        be updated, an appropriate error message is displayed.

        Args:
            None

        Returns:
            None

        Raises:
            None: All exceptions are handled internally. A ValueError is caught and an error message
            is displayed if any of the input fields are invalid, empty, or if tables and participants
            are not greater than zero.

        UI Feedback:
            - Displays a success message if the data is saved successfully.
            - Displays an error message if input validation fails or if the data cannot be saved.

        """
        date = self.dateEdit.text()
        name = self.nameEdit.text()
        address = self.addressEdit.text()
        city = self.cityEdit.text()
        plzCode = self.plzEdit.text()
        state = self.stateCombo.currentText()
        tables = self.tableEdit.text()
        participants = self.participantEdit.text()

        try:
            # Validate that all required fields are filled
            if not date or not name or not address or not city or not plzCode or not state or not tables or not participants:
                raise ValueError("All fields must be filled in.")

            # Convert tables and participants to integers and check if they are greater than zero
            tables_int = int(tables)
            participants_int = int(participants)

            if tables_int <= 0 or participants_int <= 0:
                raise ValueError("Tables and participants must be greater than zero.")

            # Get coordinates and update the Excel file
            latitude, longitude = self.get_coordinates(city, state, plzCode, address)
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
                tables_int,
                participants_int
            )
            df_events, df_stats = self.read_excel_file(self.excelFilePath)
            self.recalculateStatistics(df_events, df_stats)

            if not successFlag:
                self.create_msg_box('Input Failed!',
                                    'Data not saved.\n\nPlease check that the Excel file is not open and try again.',
                                    'warning')
            else:
                self.create_msg_box('Input Successful!',
                                    'Data saved successfully!\n\nClick on "Update Plot" to see the changes.')
                # Clear inputs
                self.clearAll()

        except ValueError as e:
            # Display an error message if any validation fails
            self.create_msg_box('Input Error', f'{str(e)}\n\nData not saved.', 'warning')

    def getPlotItem(self):
        """
       Retrieves the plot item from the DataFrame based on the currently selected item in the plot list widget.

       This method searches the `df_views` DataFrame for a row where the 'View' column matches the text of the currently
       selected item in the `plotListWidget`. If a matching entry is found, the first row of that entry is returned.

       Args:
           None

       Returns:
           pd.Series: The first row of the matching entry as a Pandas Series if found, otherwise an empty Series.
       """
        entry = self.df_views.loc[self.df_views['View'] == self.plotListWidget.currentItem().text()]
        if not entry.empty:
            entry = entry.iloc[0]  # Select the first row
        return entry

    def onPlotButtonClicked(self, doSave=False):
        """
        Handles the event of the Plot Update button being clicked.

        Reads the events and stats out of the ClimatePlotter excel document, retrieves the desired view from
        the views widget, then sends the information to the plot_map function.

        Args:
            None

        Returns:
            None

        """
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
        """
        Initialize the UI of the plot view widget on the right hand side of the program.

        Args:
            None

        Returns:
            None

        """
        self.plotListWidget.clear()
        self.plotListWidget.addItems(self.df_views['View'].tolist())
        self.plotListWidget.setCurrentRow(0)

    def onplotListWidgetClicked(self):
        """
        Handles the event of an item in the Views list being clicked upon.

        Gets the data of the selected view and updates the text fields to match.

        Args:
            None

        Returns:
            None

        """
        entry = self.getPlotItem()
        self.viewName.setText(entry['View'])
        self.viewLat.setText(str(entry['lat_0']))
        self.viewLon.setText(str(entry['lon_0']))
        self.viewLLCLat.setText(str(entry['llcrnrlat']))
        self.viewLLCLon.setText(str(entry['llcrnrlon']))
        self.viewURCLat.setText(str(entry['urcrnrlat']))
        self.viewURCLon.setText(str(entry['urcrnrlon']))

    def ClearViewText(self):
        """
        Clears the text fields associated with the Views.

        Args:
            None

        Returns:
            None

        """
        self.viewName.clear()
        self.viewLat.clear()
        self.viewLon.clear()
        self.viewLLCLat.clear()
        self.viewLLCLon.clear()
        self.viewURCLat.clear()
        self.viewURCLon.clear()

    def onPreViewButtonClicked(self):
        """
        Handles the event when the 'Pre-View' button is clicked.

        This method validates and converts the input fields required for plotting. It checks that all necessary
        fields (latitude, longitude, lower-left corner, upper-right corner coordinates, and view name) are filled
        and correctly formatted as floats. If any validation fails, an error message is displayed. If all inputs
        are valid, it proceeds to plot the map using the provided data.

        Args:
            None

        Returns:
            None

        Raises:
            None: All exceptions are handled internally. A ValueError is caught and an error message is displayed
            if any input conversion fails or if the view name is empty.

        UI Feedback:
            - Displays an error message if input validation fails, indicating the specific issue.
        """
        try:
            # Convert input text to floats
            lat = float(self.viewLat.text())
            lon = float(self.viewLon.text())
            llc_lat = float(self.viewLLCLat.text())
            llc_lon = float(self.viewLLCLon.text())
            urc_lat = float(self.viewURCLat.text())
            urc_lon = float(self.viewURCLon.text())
            view_name = self.viewName.text()

            # Check if any field is empty after conversion
            if view_name == '':
                raise ValueError("View name cannot be empty.")

            # If all conversions succeed, proceed with plotting
            df_events, df_stats = self.read_excel_file(self.excelFilePath)
            self.plot_map(
                df_events,
                df_stats,
                self.plotPath,
                self.canvas,
                lat,
                lon,
                llc_lat,
                llc_lon,
                urc_lat,
                urc_lon,
                view_name,
                doSave=False
            )

        except ValueError as e:
            # Display an error message if there's a problem with conversion or an empty field
            self.create_msg_box('Input Error', f'Invalid input: {str(e)}\n\nPlease correct the input and try again.',
                                'warning')

    def onCityLookupButtonClicked(self):
        """
        Handles the event when the 'City Lookup' button is clicked.

        This method performs a lookup for the city and state combination entered by the user. It retrieves the
        corresponding address information using the `get_address` method and populates the relevant fields
        (view name, latitude, longitude, and bounding box coordinates) in the UI. If no address is found, an
        error message is displayed. If the address is found, the method automatically triggers a preview of the
        view by calling `onPreViewButtonClicked`.

        Args:
            None

        Returns:
            None

        Raises:
            None: All exceptions are handled internally. An `AddressError` is caught and an error message is displayed
            if no address is found or if there is an issue with the address lookup.

        UI Feedback:
            - Populates the UI fields with the retrieved address information if successful.
            - Displays an error message if the address lookup fails.
        """
        try:
            raw_address = self.get_address(self.cityName.text() + ', ' + self.stateLookupCombo.currentText())
            if raw_address == None:
                raise AddressError('No address found')
            self.viewName.setText(raw_address['name'])
            self.viewLat.setText(str(raw_address['lat']))
            self.viewLon.setText(str(raw_address['lon']))
            self.viewLLCLat.setText(str(raw_address['boundingbox'][0]))
            self.viewLLCLon.setText(str(raw_address['boundingbox'][2]))
            self.viewURCLat.setText(str(raw_address['boundingbox'][1]))
            self.viewURCLon.setText(str(raw_address['boundingbox'][3]))
            self.onPreViewButtonClicked()
            # return
        except AddressError as e:
            self.create_msg_box("Address Error",
                                f"Error: {str(e)} \n\n"
                                f"Check that you have selected the correct state where the city is.",
                                'warning')
            return


    def onViewAddButtonClicked(self):
        """
        Handles the event when the 'Add View' button is clicked.

        This method validates the input fields for adding or updating a view configuration. It checks that all
        necessary fields (latitude, longitude, lower-left corner, upper-right corner coordinates, and view name)
        are filled. If any field is empty, an error message is displayed. If all inputs are valid, the method
        checks if the view name already exists in the `df_views` DataFrame:

        - If the view exists, the existing entry is updated with the new coordinates.
        - If the view does not exist, a new entry is added to the DataFrame.

        The updated DataFrame is then saved to an Excel file, and the view list widget is refreshed to reflect the changes.

        Args:
            None

        Returns:
            None

        Raises:
            None: All exceptions are handled internally. An error message is displayed if any input field is empty.

        UI Feedback:
            - Displays a success message if the view is successfully added or updated.
            - Displays an error message if any input field is empty.
        """
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
                self.df_views.loc[len(self.df_views.index) + 1] = [view_name, float(lat), float(lon), float(llc_lon),
                                                                   float(llc_lat), float(urc_lon), float(urc_lat)]
                self.df_views.to_excel(self.viewsFilePath, index=False)
                self.setupPlotListWidget()
                self.ClearViewText()
                self.create_msg_box('Input Successful!',
                                    'View saved successfully')

    def onViewRemoveButtonClicked(self):
        """
        Handles the event when the 'Remove View' button is clicked.

        This method removes the selected view from the `df_views` DataFrame, except for default views such as
        'Deutschland' and any views listed in `stateList`. If an attempt is made to remove a default view, an
        error message is displayed, and the removal is aborted. If the selected view is successfully removed,
        the DataFrame is updated, saved to an Excel file, and the plot list widget is refreshed. The input fields
        are cleared, and a success message is displayed.

        Args:
            None

        Returns:
            None

        Raises:
            None: All exceptions are handled internally. An error message is displayed if the user tries to remove
            a default view.

        UI Feedback:
            - Displays a success message if the view is removed successfully.
            - Displays an error message if the user attempts to remove a default view.
        """
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
    sys.exit(app.exec())


if __name__ == '__main__':
    run_app()
