from geopy.geocoders import Nominatim
import pandas as pd
from mpl_toolkits.basemap import Basemap
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from matplotlib import cm
import sys
import os

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
        # self.excelFilePath = os.path.join('Plotter_Output', 'ClimatePlotter.xlsx')
        self.plotPath = 'Plotter_Output'

        self.radio_buttons = []

        self.setWindowTitle("Lecture Map Plotter")
        self.setGeometry(100, 100, 1200, 900)

        self.germanyDict = {
            'Deutschland': {
                "lat_0": 51.1657,
                "lon_0": 10.4515,
                "llcrnrlon": 5.866,
                "llcrnrlat": 47.2701,
                "urcrnrlon": 15.0418,
                "urcrnrlat": 55.0583
            }
        }

        self.stateDict = {
            'Baden-Württemberg': {
                "lat_0": 48.6616,
                "lon_0": 9.3501,
                "llcrnrlon": 7.511,
                "llcrnrlat": 47.533,
                "urcrnrlon": 10.491,
                "urcrnrlat": 49.791
            },
            'Bayern': {
                "lat_0": 48.9463,
                "lon_0": 11.4039,
                "llcrnrlon": 8.979,
                "llcrnrlat": 47.270,
                "urcrnrlon": 13.839,
                "urcrnrlat": 50.564
            },
            'Berlin': {
                "lat_0": 52.5200,
                "lon_0": 13.4050,
                "llcrnrlon": 13.088,
                "llcrnrlat": 52.338,
                "urcrnrlon": 13.761,
                "urcrnrlat": 52.675
            },
            'Brandenburg': {
                "lat_0": 52.4084,
                "lon_0": 12.5625,
                "llcrnrlon": 11.267,
                "llcrnrlat": 51.359,
                "urcrnrlon": 14.765,
                "urcrnrlat": 53.558
            },
            'Bremen': {
                "lat_0": 53.0793,
                "lon_0": 8.8017,
                "llcrnrlon": 8.480,
                "llcrnrlat": 53.01,
                "urcrnrlon": 9.0,
                "urcrnrlat": 53.24
            },
            'Hamburg': {
                "lat_0": 53.5511,
                "lon_0": 9.9937,
                "llcrnrlon": 9.731,
                "llcrnrlat": 53.395,
                "urcrnrlon": 10.325,
                "urcrnrlat": 53.75
            },
            'Hessen': {
                "lat_0": 50.6521,
                "lon_0": 9.1624,
                "llcrnrlon": 7.773,
                "llcrnrlat": 49.396,
                "urcrnrlon": 10.238,
                "urcrnrlat": 51.657
            },
            'Mecklenburg-Vorpommern': {
                "lat_0": 53.6127,
                "lon_0": 12.4296,
                "llcrnrlon": 10.593,
                "llcrnrlat": 53.10,
                "urcrnrlon": 14.429,
                "urcrnrlat": 54.684
            },
            'Niedersachsen': {
                "lat_0": 52.6367,
                "lon_0": 9.8451,
                "llcrnrlon": 6.533,
                "llcrnrlat": 51.295,
                "urcrnrlon": 11.597,
                "urcrnrlat": 54.167
            },
            'Nordrhein-Westfalen': {
                "lat_0": 51.4332,
                "lon_0": 7.6616,
                "llcrnrlon": 5.866,
                "llcrnrlat": 50.322,
                "urcrnrlon": 9.461,
                "urcrnrlat": 52.530
            },
            'Rheinland-Pfalz': {
                "lat_0": 49.9825,
                "lon_0": 7.3090,
                "llcrnrlon": 6.013,
                "llcrnrlat": 48.971,
                "urcrnrlon": 8.52,
                "urcrnrlat": 50.941
            },
            'Saarland': {
                "lat_0": 49.3965,
                "lon_0": 6.9783,
                "llcrnrlon": 6.35,
                "llcrnrlat": 49.112,
                "urcrnrlon": 7.42,
                "urcrnrlat": 49.656
            },
            'Sachsen': {
                "lat_0": 51.1045,
                "lon_0": 13.2017,
                "llcrnrlon": 11.872,
                "llcrnrlat": 50.175,
                "urcrnrlon": 15.041,
                "urcrnrlat": 51.649
            },
            'Sachsen-Anhalt': {
                "lat_0": 51.9715,
                "lon_0": 11.4697,
                "llcrnrlon": 10.569,
                "llcrnrlat": 50.951,
                "urcrnrlon": 13.2,
                "urcrnrlat": 53.055
            },
            'Schleswig-Holstein': {
                "lat_0": 54.2194,
                "lon_0": 9.6961,
                "llcrnrlon": 8.120,
                "llcrnrlat": 53.359,
                "urcrnrlon": 11.43,
                "urcrnrlat": 55.058
            },
            'Thüringen': {
                "lat_0": 50.8614,
                "lon_0": 11.0522,
                "llcrnrlon": 9.872,
                "llcrnrlat": 50.204,
                "urcrnrlon": 12.652,
                "urcrnrlat": 51.637
            }
        }

        self.cityDict = {
            'Berlin': {
                "plz": [
                    '10115', '10117', '10119', '10179', '10315', '10317', '10318', '10319', '10365', '10367', '10369',
                    '10405',
                    '10407', '10409', '10435', '10437', '10439', '10551', '10553', '10555', '10557', '10559', '10585',
                    '10587',
                    '10589', '10623', '10625', '10627', '10629', '10707', '10709', '10711', '10713', '10715', '10717',
                    '10719',
                    '10777', '10779', '10781', '10783', '10785', '10787', '10789', '10823', '10825', '10827', '10829',
                    '12043',
                    '12045', '12047', '12049', '12051', '12053', '12055', '12057', '12059', '12099', '12101', '12103',
                    '12105',
                    '12107', '12109', '12157', '12159', '12161', '12163', '12165', '12167', '12169', '12203', '12205',
                    '12207',
                    '12209', '12247', '12249', '12277', '12279', '12305', '12307', '12309', '12347', '12349', '12351',
                    '12353',
                    '12355', '12357', '12359', '12435', '12437', '12439', '12459', '12487', '12489', '12524', '12526',
                    '12527',
                    '12557', '12559', '12587', '12589', '12619', '12621', '12623', '12627', '12629', '12679', '12681',
                    '12683',
                    '12685', '12687', '12689', '13051', '13053', '13055', '13057', '13059', '13086', '13088', '13089',
                    '13125',
                    '13127', '13129', '13156', '13158', '13159', '13187', '13189', '13347', '13349', '13351', '13353',
                    '13355',
                    '13357', '13359', '13403', '13405', '13407', '13409', '13435', '13437', '13439', '13465', '13467',
                    '13469',
                    '13503', '13505', '13507', '13509', '13581', '13585', '13591', '13597', '13599', '13627', '13629',
                    '14050',
                    '14052', '14053', '14055', '14057', '14059', '14089', '14109', '14129', '14131', '14163', '14165',
                    '14167',
                    '14169', '14193', '14195', '14197', '14199'
                ],
                "lat_0": 52.532,
                "lon_0": 13.3922,
                "llcrnrlon": 13.1055859,
                "llcrnrlat": 52.3932541,
                "urcrnrlon": 13.6266696,
                "urcrnrlat": 52.6131758
            },
            'Hamburg': {
                "plz": [
                    '20095', '20097', '20099', '20144', '20146', '20148', '20149', '20249', '20251', '20253', '20255',
                    '20257',
                    '20259', '20354', '20355', '20357', '20359', '20457', '20459', '20535', '20537', '20539', '21029',
                    '21031',
                    '21033', '21035', '21037', '21039', '21073', '21075', '21077', '21079', '21107', '21109', '21129',
                    '21147',
                    '21149', '22041', '22043', '22045', '22047', '22049', '22081', '22083', '22085', '22087', '22089',
                    '22111',
                    '22113', '22115', '22117', '22119', '22143', '22145', '22147', '22149', '22159', '22175', '22177',
                    '22179',
                    '22297', '22299', '22301', '22303', '22305', '22307', '22309', '22335', '22337', '22339', '22359',
                    '22391',
                    '22393', '22395', '22397', '22399', '22415', '22417', '22419', '22453', '22455', '22457', '22459',
                    '22523',
                    '22525', '22527', '22529', '22547', '22549', '22559', '22587', '22589', '22605', '22607', '22609',
                    '22761',
                    '22763', '22765', '22767', '22769'
                ],
                "lat_0": 53.5544,
                "lon_0": 9.9946,
                "llcrnrlon": 9.8073217,
                "llcrnrlat": 53.4457532,
                "urcrnrlon": 10.2137762,
                "urcrnrlat": 53.6817768
            },
            'München': {
                "plz": [
                    '80331', '80333', '80335', '80336', '80337', '80339', '80469', '80538', '80539', '80634', '80636',
                    '80637',
                    '80638', '80639', '80686', '80687', '80689', '80796', '80797', '80798', '80799', '80801', '80802',
                    '80803',
                    '80804', '80805', '80807', '80809', '80933', '80935', '80937', '80939', '80992', '80993', '80995',
                    '80997',
                    '80999', '81241', '81243', '81245', '81247', '81369', '81371', '81373', '81375', '81377', '81379',
                    '81475',
                    '81476', '81477', '81479', '81539', '81541', '81543', '81545', '81547', '81549', '81667', '81669',
                    '81671',
                    '81673', '81675', '81677', '81679', '81735', '81737', '81739', '81825', '81827', '81829', '81925',
                    '81927',
                    '81929'
                ],
                "lat_0": 48.1351,
                "lon_0": 11.5820,
                "llcrnrlon": 11.360,
                "llcrnrlat": 48.061,
                "urcrnrlon": 11.722,
                "urcrnrlat": 48.248
            },
            'Köln': {
                "plz": [
                    '50667', '50668', '50670', '50672', '50674', '50676', '50677', '50678', '50679', '50733', '50735',
                    '50737',
                    '50739', '50765', '50767', '50769', '50823', '50825', '50827', '50829', '50858', '50859', '50931',
                    '50933',
                    '50935', '50937', '50939', '50968', '50969', '50996', '50997', '50999', '51061', '51063', '51065',
                    '51067',
                    '51069', '51103', '51105', '51107', '51109', '51143', '51145', '51147', '51149', '51467'
                ],
                "lat_0": 50.9384,
                "lon_0": 6.9543,
                "llcrnrlon": 6.8824458,
                "llcrnrlat": 50.8910803,
                "urcrnrlon": 7.0653033,
                "urcrnrlat": 51.0145577
            },
            'Frankfurt': {
                "plz": [
                    '60308', '60311', '60313', '60314', '60316', '60318', '60320', '60322', '60323', '60325', '60326',
                    '60327',
                    '60329', '60385', '60386', '60388', '60389', '60431', '60433', '60435', '60437', '60438', '60439',
                    '60486',
                    '60487', '60488', '60489', '60528', '60529', '60549', '60594', '60596', '60598', '60599', '65929',
                    '65931',
                    '65933', '65934', '65936'
                ],
                "lat_0": 50.1167,
                "lon_0": 8.6833,
                "llcrnrlon": 8.4888008,
                "llcrnrlat": 50.0133053,
                "urcrnrlon": 8.7834831,
                "urcrnrlat": 50.2164379
            },
            'Stuttgart': {
                "plz": [
                    '70173', '70174', '70176', '70178', '70180', '70182', '70184', '70186', '70188', '70190', '70191',
                    '70192',
                    '70193', '70195', '70197', '70199', '70327', '70329', '70372', '70374', '70376', '70378', '70435',
                    '70437',
                    '70439', '70469', '70499', '70563', '70565', '70567', '70569', '70597', '70599', '70619', '70629'
                ],
                "lat_0": 48.7667,
                "lon_0": 9.1833,
                "llcrnrlon": 9.1343004,
                "llcrnrlat": 48.7545509,
                "urcrnrlon": 9.2414900,
                "urcrnrlat": 48.8204811
            },
            'Düsseldorf': {
                "plz": [
                    '40210', '40211', '40212', '40213', '40215', '40217', '40219', '40221', '40223', '40225', '40227',
                    '40229',
                    '40231', '40233', '40235', '40237', '40239', '40468', '40470', '40472', '40474', '40476', '40477',
                    '40479',
                    '40489', '40545', '40547', '40549', '40589', '40591', '40593', '40595', '40597', '40599', '40625',
                    '40627',
                    '40629'
                ],
                "lat_0": 51.2216,
                "lon_0": 6.7898,
                "llcrnrlon": 6.6985260,
                "llcrnrlat": 51.1237393,
                "urcrnrlon": 6.9201588,
                "urcrnrlat": 51.3401162
            },
            'Dortmund': {
                "plz": [
                    '44135', '44137', '44139', '44141', '44143', '44145', '44147', '44149', '44225', '44227', '44229',
                    '44263',
                    '44265', '44267', '44269', '44287', '44289', '44309', '44319', '44328', '44329', '44339', '44357',
                    '44359',
                    '44369', '44379', '44388'
                ],
                "lat_0": 51.5125,
                "lon_0": 7.477,
                "llcrnrlon": 7.3189002,
                "llcrnrlat": 51.4203522,
                "urcrnrlon": 7.6285024,
                "urcrnrlat": 51.5902729
            },
            'Essen': {
                "plz": [
                    '45127', '45128', '45130', '45131', '45133', '45134', '45136', '45138', '45139', '45141', '45143',
                    '45144',
                    '45145', '45147', '45149', '45219', '45239', '45257', '45259', '45276', '45277', '45279', '45289',
                    '45307',
                    '45309', '45326', '45327', '45329', '45355', '45356', '45357', '45359'
                ],
                "lat_0": 51.4536,
                "lon_0": 7.0102,
                "llcrnrlon": 6.9009234,
                "llcrnrlat": 51.3477941,
                "urcrnrlon": 7.1307693,
                "urcrnrlat": 51.5340233
            },
            'Leipzig': {
                "plz": [
                    '04109'
                ],
                "lat_0": 51.342,
                "lon_0": 12.375,
                "llcrnrlon": 12.2465782,
                "llcrnrlat": 51.2638967,
                "urcrnrlon": 12.5297905,
                "urcrnrlat": 51.4002769
            },
            'Bremen': {
                "plz": [
                    '28195', '28197', '28199', '28201', '28203', '28205', '28207', '28209', '28211', '28213', '28215',
                    '28217',
                    '28219', '28237', '28239', '28259', '28277', '28279', '28307', '28309', '28325', '28327', '28329',
                    '28355',
                    '28357', '28359', '28717', '28719', '28755', '28757', '28759', '28777', '28779'
                ],
                "lat_0": 53.0889,
                "lon_0": 8.7906,
                "llcrnrlon": 8.4947922,
                "llcrnrlat": 53.0221694,
                "urcrnrlon": 8.9870293,
                "urcrnrlat": 53.2204517
            },
            'Dresden': {
                "plz": [
                    '01067', '01069', '01097', '01099', '01108', '01109', '01127', '01129', '01139', '01156', '01157',
                    '01159',
                    '01169', '01187', '01189', '01217', '01219', '01237', '01239', '01257', '01259', '01277', '01279',
                    '01307',
                    '01309', '01324', '01326', '01328', '01462', '01465', '01478'
                ],
                "lat_0": 51.05,
                "lon_0": 13.75,
                "llcrnrlon": 13.5754990,
                "llcrnrlat": 50.9758015,
                "urcrnrlon": 13.9326637,
                "urcrnrlat": 51.1658162
            },
            'Hanover': {
                "plz": [
                    '30159', '30161', '30163', '30165', '30167', '30169', '30171', '30173', '30175', '30177', '30179',
                    '30419',
                    '30449', '30451', '30453', '30455', '30457', '30459', '30519', '30521', '30539', '30559', '30625',
                    '30627',
                    '30629', '30655', '30657', '30659', '30669'
                ],
                "lat_0": 52.3736,
                "lon_0": 9.7371,
                "llcrnrlon": 9.6256456,
                "llcrnrlat": 52.3076533,
                "urcrnrlon": 9.9171856,
                "urcrnrlat": 52.4484497
            },
            'Nürnberg': {
                "plz": [
                    '90402', '90403', '90408', '90409', '90411', '90419', '90425', '90427', '90429', '90431', '90439',
                    '90441',
                    '90443', '90449', '90451', '90453', '90455', '90459', '90461', '90469', '90471', '90473', '90475',
                    '90478',
                    '90480', '90482', '90489', '90491'
                ],
                "lat_0": 49.4504,
                "lon_0": 11.0778,
                "llcrnrlon": 10.9940730,
                "llcrnrlat": 49.3428032,
                "urcrnrlon": 11.1882136,
                "urcrnrlat": 49.5367797
            },
            'Duisberg': {
                "plz": [
                    '47198'
                ],
                "lat_0": 51.4499,
                "lon_0": 6.6901,
                "llcrnrlon": 6.6300531,
                "llcrnrlat": 51.3300235,
                "urcrnrlon": 6.8401972,
                "urcrnrlat": 51.5621746
            },
            'Bochum': {
                "plz": [
                    '44787', '44789', '44791', '44793', '44795', '44797', '44799', '44801', '44803', '44805', '44807',
                    '44809',
                    '44866', '44867', '44869', '44879', '44892', '44894'
                ],
                "lat_0": 51.4803,
                "lon_0": 7.2183,
                "llcrnrlon": 7.1182351,
                "llcrnrlat": 51.4062166,
                "urcrnrlon": 7.3439578,
                "urcrnrlat": 51.5315915
            },
            'Wuppertal': {
                "plz": [
                    '42103', '42105', '42107', '42109', '42111', '42113', '42115', '42117', '42119', '42275', '42277',
                    '42279',
                    '42281', '42283', '42285', '42287', '42289', '42327', '42329', '42349', '42369', '42389', '42399'
                ],
                "lat_0": 51.2569,
                "lon_0": 7.1505,
                "llcrnrlon": 7.0911096,
                "llcrnrlat": 51.2202031,
                "urcrnrlon": 7.2393535,
                "urcrnrlat": 51.2977357
            },
            'Bielefeld': {
                "plz": [
                    '33602', '33604', '33605', '33607', '33609', '33611', '33613', '33615', '33617', '33619', '33647',
                    '33649',
                    '33659', '33689', '33699', '33719', '33729', '33739'
                ],
                "lat_0": 52.0245,
                "lon_0": 8.5326,
                "llcrnrlon": 8.4680291,
                "llcrnrlat": 51.9735205,
                "urcrnrlon": 8.6149822,
                "urcrnrlat": 52.0601791
            },
            'Bonn': {
                "plz": [
                    '53111', '53113', '53115', '53117', '53119', '53121', '53123', '53125', '53127', '53129', '53173',
                    '53175',
                    '53177', '53179', '53225', '53227', '53229'
                ],
                "lat_0": 50.7362,
                "lon_0": 7.1002,
                "llcrnrlon": 7.0315679,
                "llcrnrlat": 50.6815715,
                "urcrnrlon": 7.1774806,
                "urcrnrlat": 50.7640391
            },
            'Münster': {
                "plz": [
                    48079
                ],
                "lat_0": 51.9625,
                "lon_0": 7.6256,
                "llcrnrlon": 7.480,
                "llcrnrlat": 51.892,
                "urcrnrlon": 7.770,
                "urcrnrlat": 52.042
            },
            'Karlsruhe': {
                "plz": [
                    '76131', '76133', '76135', '76137', '76139', '76149', '76185', '76187', '76189', '76199', '76227',
                    '76228',
                    '76229'
                ],
                "lat_0": 49.0069,
                "lon_0": 8.4037,
                "llcrnrlon": 8.282,
                "llcrnrlat": 48.935,
                "urcrnrlon": 8.524,
                "urcrnrlat": 49.073
            }
        }

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
        location = geolocator.geocode(address + ", " + city_name + ", " + state_name + ", " + plz_code + ", Germany")
        if location:
            return location.latitude, location.longitude
        else:
            return None, None

    def read_excel_file(self, file_path):
        try:
            df = pd.read_excel(file_path)
            return df
        except pd.errors.EmptyDataError:
            df = pd.DataFrame(
                columns=['Name', 'City', 'State', 'PLZ', 'Latitude', 'Longitude', 'EventCount', 'CityEventTotal',
                         'TotalTables', 'TotalParticipants'])
            return df

    def update_excel(self, file_path, name, address, city, state, plz, lat, lon, tables=0,
                     participants=0):
        df = pd.read_excel(file_path)
        school_exists = ((df['Name'] == name) & (df['City'] == city) & (df['State'] == state)).any().any()
        if school_exists:
            df.loc[(df['Name'] == name) & (df['City'] == city) & (df['State'] == state), 'EventCount'] += 1
            df.loc[(df['Name'] == name) & (df['City'] == city) & (df['State'] == state), 'TotalTables'] += int(tables)
            df.loc[(df['Name'] == name) & (df['City'] == city) & (df['State'] == state), 'TotalParticipants'] += int(
                participants)
        else:
            df.loc[len(df.index) + 1] = [name, address, city, state, plz, lat, lon, 1, -1, int(tables),
                                         int(participants)]
        city_event_total = df[df['City'] == city]['EventCount'].sum()
        df.loc[df['City'] == city, 'CityEventTotal'] = int(city_event_total)
        df.to_excel(file_path, index=False)
        return df

    def plot_map(self, df, save_path, canvas, lat, lon, llc_lat, llc_lon, urc_lat, urc_lon, view='Deutschland'):
        canvas.figure.clf()
        ax = canvas.figure.add_subplot(111)
        '''
        resolution: c (crude), l (low), i (intermediate), h (high), f (full) or None
        projection: 'merc' (Mercator), 'cyl' (Cylindrical Equidistant), 'mill' (Miller Cylindrical), 'gall' (Gall Stereographic Cylindrical), 'cea' (Cylindrical Equal Area), 'lcc' (Lambert Conformal), 'tmerc' (Transverse Mercator), 'omerc' (Oblique Mercator), 'nplaea' (North-Polar Lambert Azimuthal), 'npaeqd' (North-Polar Azimuthal Equidistant), 'nplaea' (South-Polar Lambert Azimuthal), 'spaeqd' (South-Polar Azimuthal Equidistant), 'aea' (Albers Equal Area), 'stere' (Stereographic), 'robin' (Robinson), 'eck4' (Eckert IV), 'eck6' (Eckert VI), 'kav7' (Kavrayskiy VII), 'mbtfpq' (McBryde-Thomas Flat-Polar Quartic), 'sinu' (Sinusoidal), 'gall' (Gall Stereographic Cylindrical), 'hammer' (Hammer), 'moll' (Mollweid
        espg: 3857 (Web Mercator)
        '''
        m = Basemap(resolution='h', lat_0=lat, lon_0=lon, llcrnrlon=llc_lon, llcrnrlat=llc_lat,
                    urcrnrlon=urc_lon, urcrnrlat=urc_lat, epsg=3857)
        m.drawcountries()

        '''
        https://gdz.bkg.bund.de/index.php/default/webdienste.html
        https://gdz.bkg.bund.de/index.php/default/webdienste/basemap-webdienste/wms-basemapde-webraster-wms-basemapde-webraster.html
        de_basemapde_web_raster_farbe
        de_basemapde_web_raster_grau
        '''
        # wms_server = 'https://sgx.geodatenzentrum.de/wms_basemapde?SERVICE=WMS&Request=GetCapabilities'
        # m.wmsimage(wms_server, layers=["de_basemapde_web_raster_farbe"], verbose=True)

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
        m.wmsimage(wms_server, layers=["web"], verbose=True)

        m.drawcoastlines()
        if view in self.cityDict:
            df1 = df.loc[df['City'] == view].reset_index(drop=True)
            df_max = max(df1['TotalParticipants'])
            for k, d in df1.iterrows():
                msize = d['TotalParticipants'] * 5
                if msize > 200:
                    msize = 200
                col = cm.gist_rainbow(d['TotalParticipants'] / df_max)
                x, y = m(d['Longitude'], d['Latitude'])
                ax.scatter(x, y, label=k, s=msize, color=col)
        else:
            '''
            Draw the German States
            '''
            shapePath = os.path.join(os.path.dirname(__file__), 'shapefiles', 'DEU_adm1')
            m.readshapefile(shapePath, 'areas')
            # m.readshapefile('./shapefiles/DEU_adm1', 'areas')
            df_poly = pd.DataFrame(columns=['shapes', 'area'])
            for info, shape in zip(m.areas_info, m.areas):
                shape_array = np.array(shape)
                df_poly.loc[len(df_poly.index) + 1] = [Polygon(shape_array), info['NAME_1']]

            df1 = df.groupby('CityEventTotal')
            colors = iter(cm.gnuplot(np.linspace(0, 1, len(df1.groups))))

            for group_cluster, data in df1:
                msize = group_cluster * 10
                if msize > 50:
                    msize = 50
                x, y = m(data['Longitude'], data['Latitude'])
                ax.scatter(x, y, label=group_cluster, s=msize, color=next(colors))
        save_path = os.path.join(os.path.dirname(__file__), save_path, f'{pd.Timestamp.now().strftime("%Y-%m-%d_H%HM%MS%S")}_{view}.png')
        # save_path = save_path + f'{pd.Timestamp.now().strftime("%Y-%m-%d_%H-%M-%S")}_{view}.png'
        plt.savefig(save_path, format='png', dpi=300)
        canvas.draw()


    def drawInitialMap(self):
        df_fresks = self.read_excel_file(self.excelFilePath)
        self.plot_map(df_fresks,
                 self.plotPath,
                 self.canvas,
                 self.germanyDict['Deutschland']['lat_0'],
                 self.germanyDict['Deutschland']['lon_0'],
                 self.germanyDict['Deutschland']['llcrnrlat'],
                 self.germanyDict['Deutschland']['llcrnrlon'],
                 self.germanyDict['Deutschland']['urcrnrlat'],
                 self.germanyDict['Deutschland']['urcrnrlon']
                 )

    def addExcelFiles(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        dialog.setNameFilter("Excel files (*.xlsx)")
        if dialog.exec():
            fileNames = dialog.selectedFiles()
            for fileName in fileNames:
                self.bulkImportList.addItem(fileName)

    def removeExcelFiles(self):
        self.bulkImportList.takeItem(self.bulkImportList.currentRow())

    def clearExcelFiles(self):
        self.bulkImportList.clear()

    def bulkImportExcelFiles(self):
        for i in reversed(range(self.bulkImportList.count())):
            bulkFilePath = self.bulkImportList.item(i).text()
            df = pd.read_excel(bulkFilePath)
            df = df.fillna('')  # Replace NaN with empty string
            for index, row in df.iterrows():
                date = row['Datum']
                time = row['Uhrzeit']
                name = row['Schul/Uni']
                address = str(row['Adresse'])
                city = str(row['Stadt'])
                state = str(row['Bundesland'])
                plzCode = str(row['PLZ'])
                table = str(row['Tische'])
                participant = str(row['Teilnehmer'])
                latitude, longitude = self.get_coordinates(address, city, state, plzCode)
                self.update_excel(self.excelFilePath, name, address, city, state, plzCode, latitude, longitude, table, participant)
            if os.path.exists(bulkFilePath):
                os.remove(bulkFilePath)
            self.bulkImportList.takeItem(i)
        self.drawInitialMap()

    def onLookupAddressButtonClicked(self):
        name = self.nameEdit.text()
        raw_address = self.get_address(name).split(', ')
        self.nameEdit.setText(raw_address[0])
        self.plzEdit.setText(raw_address[-2])
        self.cityEdit.setText(raw_address[-4])
        state = raw_address[-3]
        i = -1
        for item in self.stateDict:
            i += 1
            if state == item:
                self.stateCombo.setCurrentIndex(i)
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
        self.stateCombo.setCurrentIndex(0)
        self.tableEdit.setValue(0)
        self.participantEdit.setValue(0)

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
        BulkImportBox = QGridLayout()
        InputBox = QGridLayout()
        middleBox = QVBoxLayout()
        rightBox = QVBoxLayout()

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

        self.bulkImportButton = QPushButton("Bulk Import", self)

        BulkImportBox.addWidget(self.bulkImportList, 0, 0, 3, 1)
        BulkImportBox.addWidget(self.addButton, 0, 1)
        BulkImportBox.addWidget(self.removeButton, 1, 1)
        BulkImportBox.addWidget(self.clearButton, 2, 1)
        BulkImportBox.addWidget(self.bulkImportButton, 3, 0, 1, 2)

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
        for state in self.stateDict:
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
        self.plotButton.clicked.connect(self.onPlotButtonClicked)

        self.clearAllButton = QPushButton("Clear All", self)
        self.clearAllButton.clicked.connect(self.clearAll)

        InputBox.addWidget(self.nameLabel, 0, 0)
        InputBox.addWidget(self.nameEdit, 0, 1)
        InputBox.addWidget(self.addressLookupButton, 1, 0, 1, 2)
        InputBox.addWidget(self.addressLabel, 2, 0)
        InputBox.addWidget(self.addressEdit, 2, 1)
        InputBox.addWidget(self.cityLabel, 3, 0)
        InputBox.addWidget(self.cityEdit, 3, 1)
        InputBox.addWidget(self.plzLabel, 4, 0)
        InputBox.addWidget(self.plzEdit, 4, 1)
        InputBox.addWidget(self.stateLabel, 5, 0)
        InputBox.addWidget(self.stateCombo, 5, 1)
        InputBox.addWidget(self.dateLabel, 6, 0)
        InputBox.addWidget(self.dateEdit, 6, 1)
        InputBox.addWidget(self.tableLabel, 7, 0)
        InputBox.addWidget(self.tableEdit, 7, 1)
        InputBox.addWidget(self.participantLabel, 8, 0)
        InputBox.addWidget(self.participantEdit, 8, 1)
        InputBox.addWidget(self.updateCsvButton, 9, 0)
        InputBox.addWidget(self.plotButton, 9, 1)
        InputBox.addWidget(self.clearAllButton, 10, 0, 1, 2)

        '''
        middleBox
        '''
        self.figure = plt.figure()
        self.canvas = FigureCanvas(self.figure)
        middleBox.addWidget(self.canvas)  # Add the canvas to the layout

        '''
        rightBox
        '''
        for country in self.germanyDict:
            radioButton = QRadioButton(country, self)
            rightBox.addWidget(radioButton)
            self.radio_buttons.append(radioButton)
        radioButton.setChecked(True)

        rightBox.addWidget(QFrame(self, frameShape=QFrame.Shape.HLine))

        for state in self.stateDict:
            radioButton = QRadioButton(state, self)
            rightBox.addWidget(radioButton)
            self.radio_buttons.append(radioButton)

        rightBox.addWidget(QFrame(self, frameShape=QFrame.Shape.HLine))

        for city in self.cityDict:
            radioButton = QRadioButton(city, self)
            rightBox.addWidget(radioButton)
            self.radio_buttons.append(radioButton)

        leftBox.addLayout(BulkImportBox, 0, 0, 2, 2)
        leftBox.addWidget(QFrame(self, frameShape=QFrame.Shape.HLine), 3, 0, 1, 2)
        leftBox.addLayout(InputBox, 4, 0, 3, 2)

        mylayout.addLayout(leftBox)
        mylayout.addLayout(middleBox)
        mylayout.addLayout(rightBox)

        centralWidget.setLayout(mylayout)

        # Draw the initial map
        self.drawInitialMap()

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
            msgBox = QMessageBox(self)
            msgBox.setIcon(QMessageBox.Icon.Information)
            msgBox.setWindowTitle("Input Empty!")
            msgBox.setText("Fill in all the fields!\n\nData not saved.")
            msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
            msgBox.exec()
        else:
            latitude, longitude = self.get_coordinates(address, city, state, plzCode)
            self.update_excel(self.excelFilePath, name, address, city, state, plzCode, latitude, longitude, tables, participants)
            msgBox = QMessageBox(self)
            msgBox.setIcon(QMessageBox.Icon.Information)
            msgBox.setWindowTitle("Input Successful!")
            msgBox.setText("Data saved successfully!\n\nClick on 'Update Plot' to see the changes.")
            msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
            msgBox.exec()

            # Clear inputs
            self.nameEdit.clear()
            self.cityEdit.clear()
            self.addressEdit.clear()
            self.plzEdit.clear()
            self.tableEdit.setValue(0)
            self.participantEdit.setValue(0)

    def getRadioButtonIndex(self):
        for btn in self.radio_buttons:
            if btn.isChecked():
                return self.radio_buttons.index(btn)

    def getPlotItem(self):
        radioButtonIndex = self.getRadioButtonIndex()
        if radioButtonIndex == 0:
            return list(self.germanyDict.values())[radioButtonIndex], list(self.germanyDict.keys())[radioButtonIndex]
        elif 0 < radioButtonIndex < 17:
            radioButtonIndex = radioButtonIndex - 1  # because first button is for Germany, separate dict
            return list(self.stateDict.values())[radioButtonIndex], list(self.stateDict.keys())[radioButtonIndex]
        else:
            radioButtonIndex = radioButtonIndex - 17  # because first 17 buttons are for Germany and states, separate dict
            return list(self.cityDict.values())[radioButtonIndex], list(self.cityDict.keys())[radioButtonIndex]

    def onPlotButtonClicked(self):
        df_fresks = self.read_excel_file(self.excelFilePath)
        plotItem, plotItemName = self.getPlotItem()
        self.plot_map(df_fresks, self.plotPath, self.canvas, plotItem['lat_0'], plotItem['lon_0'], plotItem['llcrnrlat'],
                 plotItem['llcrnrlon'], plotItem['urcrnrlat'], plotItem['urcrnrlon'], plotItemName)

        # # Clear inputs
        # self.nameEdit.clear()
        # self.cityEdit.clear()


def run_app():
    app = QApplication([])
    mainWin = LectureMapApp()
    mainWin.show()
    # app.exec()
    sys.exit(app.exec())


if __name__ == '__main__':
    run_app()
