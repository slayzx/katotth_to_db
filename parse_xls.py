from openpyxl import load_workbook
import psycopg2
from dotenv import load_dotenv
from dotenv import dotenv_values
from icecream import ic

load_dotenv()

config = dotenv_values(".env")

print('Put the Exel file in root directory, and enter it`s name with document type. For example: katotth.xlsx')

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

past = 0

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
                district SMALLINT, 
                district_name VARCHAR(100),
                hromada SMALLINT, 
                municipal SMALLINT, 
                district_city SMALLINT, 
                division_num INT, 
                division_type VARCHAR(50), 
                division_name VARCHAR(50), 
                division_full_name VARCHAR(100) 
                );"""

drop_table = """DROP TABLE katotth;"""

workbook = load_workbook("katotth_orig.xlsx")

sheet = workbook.active
print("Specify the name (in uppercase) of the upper left and lower right cells of the table without extra rows, "
      "only data.")
# upper_left = input('upper left: ')
# lower_right = input('lower right: ')

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

        # for cell in sheet[f"{upper_left.upper()}:{lower_right.upper()}"]:  # A4 G31761
        # for cell in sheet[upper_left:lower_right]:
        for cell in sheet['A4':'G2000']:
            division_type, division_name = cell[5].value, cell[6].value
            if division_type in ['O', 'P', 'H', 'B']:
                division_full_name = f"{division_name} {obj_decode[division_type]}"
            else:
                division_full_name = f"{obj_decode[division_type]} {division_name}"

            atu_num = cell[obj_to_column[division_type]].value

            past = district

            region = int(atu_num[2:4])
            district = int(atu_num[4:6])
            hromada = int(atu_num[6:9])
            municipal = int(atu_num[9:12])
            district_city = int(atu_num[12:14])
            division_num = int(atu_num[-5:])

            if past != district:
                district_name = division_full_name

            # If object of administrative division is region or city with special status
            if division_type in ['O', 'K']:
                curs.execute(
                    'INSERT INTO katotth(region, division_num, division_type, division_name, division_full_name) '
                    'VALUES(%s, %s, %s, %s, %s)',
                    (region, division_num, obj_decode[division_type], division_name, division_full_name))
            # If object of administrative division is district
            elif division_type == 'P':
                curs.execute('''INSERT INTO katotth(region, district, district_name, division_num, division_type, division_name, division_full_name) 
                                VALUES(%s, %s, %s, %s, %s, %s, %s)''',
                             (region, district, district_name, division_num, obj_decode[division_type], division_name,
                              division_full_name))
            # If object of administrative division is territorial hromada
            elif division_type == 'H':
                curs.execute('''INSERT INTO katotth(region, district, district_name, hromada, division_num, division_type, division_name, division_full_name) 
                                VALUES(%s, %s, %s, %s, %s, %s, %s, %s)''',
                             (region, district, district_name, hromada, division_num, obj_decode[division_type], division_name,
                              division_full_name))
            # If object of administrative division is one of the city or or village or special status
            elif division_type in ['M', 'T', 'C', 'X']:
                curs.execute('''INSERT INTO katotth(region, district, district_name, hromada, municipal, division_num, division_type, division_name, division_full_name) 
                                VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s)''',
                             (region, district, district_name, hromada, municipal, division_num, obj_decode[division_type],
                              division_name, division_full_name))
            # If object of administrative division is district in city
            elif division_type == 'B':
                curs.execute('''INSERT INTO katotth(region, district, district_name, hromada, municipal, district_city, division_num, 
                                                    division_type, division_name, division_full_name) 
                                VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)''',
                             (region, district, district_name, hromada, municipal, district_city, division_num,
                              obj_decode[division_type], division_name,
                              division_full_name))
    conn.commit()

workbook.close()

'''Columns: A - region, B - district, C - hromada, D - communities, 
E - districts in cities, F - object category, G - name'''

# [cell.value for cell in sheet['A']]  пробегает все значения в колонке А
# [cell[6].value for cell in sheet.rows]  пробегает по значениям ячеек конкретной колонки
