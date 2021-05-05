from openpyxl import load_workbook
import psycopg2
from dotenv import load_dotenv
from dotenv import dotenv_values
import re

load_dotenv()

config = dotenv_values(".env")

crimea = re.compile('республіка крим', re.I)

obj_decode = {
    'O': 'область',
    'K': 'місто',  # city with special status
    'P': 'район',
    'H': 'територіальна громада',
    'M': 'місто',
    'T': 'смт',
    'C': 'село',
    'X': 'селище',
    'B': 'район у місті'
}
obj_to_column = {
    'O': 0,
    'K': 0,  # city with special status
    'P': 1,
    'H': 2,
    'M': 3,
    'T': 3,
    'C': 3,
    'X': 3,
    'B': 4
}

# variables for database: reg - region, dist - district, hrom - hromada, munic - municipalities,
#  division_name - unique number of administrative division object, div_type - type of division object
# respectively to katotth, division_name - name of administrative division object,
# division_full_name - humanised division_name + div_type

past_district = 0

region, district, district_name, hromada, municipal, district_city, division_num, division_type, division_name, division_full_name = \
    0, 0, '', 0, 0, 0, 0, '', '', ''

conn = psycopg2.connect(
    host=config['PGSQL_HOST'],
    database=config['PGSQL_DB_NAME'],
    user=config['PGSQL_DB_USER'],
    password='PGSQL_DB_PASSWD')

create_table = """create table katotth( 
                id SERIAL PRIMARY KEY, 
                region SMALLINT, 
                region_name VARCHAR(50),
                distr SMALLINT, 
                distr_name VARCHAR(50),
                hrom SMALLINT, 
                hrom_name VARCHAR(50),
                municip SMALLINT, 
                municip_name VARCHAR(50),
                distr_city SMALLINT, 
                div_num INT, 
                div_type VARCHAR(50), 
                div_name VARCHAR(50), 
                div_full_name VARCHAR(100) 
                );"""

drop_table = """DROP TABLE katotth;"""

workbook = load_workbook("katotth_orig.xlsx")

print('Put the Exel file in root directory, and enter it`s name with document type. For example: katotth.xlsx')
sheet = workbook.active
print("Specify the name (in uppercase) of the upper left and lower right cells of the table without extra rows, "
      "only data.")
upper_left = input('upper left: ')
lower_right = input('lower right: ')

try:
    with conn:
        with conn.cursor() as curs:
            curs.execute(create_table)

except psycopg2.errors.DuplicateTable:
    with conn:
        with conn.cursor() as curs:
            curs.execute(drop_table)
            curs.execute(create_table)

conn.commit()

with conn:
    with conn.cursor() as curs:

        # A4 G31761
        for cell in sheet[upper_left:lower_right]:
            division_type, division_name = cell[5].value, cell[6].value
            if division_type == 'O':
                if crimea.search(division_name):
                    division_full_name = division_name
                else:
                    division_full_name = f"{division_name} {obj_decode[division_type]}"
            elif division_type in ['P', 'H', 'B']:
                division_full_name = f"{division_name} {obj_decode[division_type]}"
            else:
                division_full_name = f"{obj_decode[division_type]} {division_name}"

            past_region, past_district, past_hromada = region, district, hromada

            atu_num = cell[obj_to_column[division_type]].value

            region = int(atu_num[2:4])
            district = int(atu_num[4:6])
            hromada = int(atu_num[6:9])
            municipal = int(atu_num[9:12])
            district_city = int(atu_num[12:14])
            division_num = int(atu_num[-5:])

            # If object of administrative division is region or city with special status
            if division_type in ['O', 'K']:
                if past_region != region:
                    region_name = division_full_name
                curs.execute(
                    """INSERT INTO katotth(region, region_name, distr, hrom, municip, distr_city, div_num, div_type,
                                           div_name, div_full_name) 
                       VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                    (region, region_name, district, hromada, municipal, district_city, division_num,
                     obj_decode[division_type], division_name, division_full_name))
            # If object of administrative division is district
            elif division_type == 'P':
                if past_district != district:
                    district_name = division_full_name
                curs.execute(
                    """INSERT INTO katotth(region, region_name, distr, distr_name, hrom, municip, distr_city, div_num,
                                           div_type, div_name, div_full_name) 
                       VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                    (region, region_name, district, district_name, hromada, municipal, district_city, division_num,
                     obj_decode[division_type], division_name, division_full_name))
            # If object of administrative division is territorial hromada
            elif division_type == 'H':
                if past_hromada != hromada:
                    hromada_name = division_full_name
                curs.execute(
                    """INSERT INTO katotth(region, region_name, distr, distr_name, hrom, hrom_name, municip, 
                                           distr_city, div_num, div_type, div_name, div_full_name) 
                       VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                    (region, region_name, district, district_name, hromada, hromada_name, municipal,
                     district_city, division_num, obj_decode[division_type], division_name, division_full_name))
            # If object of administrative division is one of the city or or village or special status
            elif division_type in ['M', 'T', 'C', 'X']:
                past_municipal = division_name
                curs.execute(
                    """INSERT INTO katotth(region, region_name, distr, distr_name, hrom, hrom_name, municip, 
                                           distr_city, div_num, div_type, div_name, div_full_name) 
                       VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                    (region, region_name, district, district_name, hromada, hromada_name, municipal,
                     district_city, division_num, obj_decode[division_type], division_name, division_full_name))
            # If object of administrative division is district in city
            elif division_type == 'B':
                curs.execute(
                    """INSERT INTO katotth(region, region_name, distr, distr_name, hrom, hrom_name, municip, 
                    municip_name, distr_city, div_num, div_type, div_name, div_full_name) 
                        VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
                    (region, region_name, district, district_name, hromada, hromada_name, municipal, past_municipal,
                     district_city, division_num, obj_decode[division_type], division_name, division_full_name))
    conn.commit()
    pass

workbook.close()

'''Columns: A - region, B - district, C - hromada, D - communities, 
E - districts in cities, F - object category, G - name'''