from openpyxl import load_workbook
import psycopg2

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
#  div_num - unique number of administrative division object, div_type - type of division object
# respectively to katotth, div_name - name of administrative division object,
# div_full_name - humanised div_name + div_type

reg, distr, hrom, munic, distr_city, div_num, div_type, div_name, div_full_name = 0, 0, 0, 0, 0, 0, '', '', ''

conn = psycopg2.connect(
    host="localhost",
    database="osm",
    user="osm",
    password="osm")

create_table = """create table katotth( 
                id SERIAL PRIMARY KEY, 
                reg SMALLINT, 
                distr SMALLINT, 
                hrom SMALLINT, 
                munic SMALLINT, 
                distr_city SMALLINT, 
                div_num INT, 
                div_type VARCHAR(50), 
                div_name VARCHAR(50), 
                div_full_name VARCHAR(100) 
                );"""
drop_table = """DROP TABLE katotth;"""


workbook = load_workbook("katotth_orig.xlsx")

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


with conn.cursor() as curs:

    # for cell in sheet[f"{upper_left.upper()}:{lower_right.upper()}"]:  # A4 G31761
    for cell in sheet['A4':'G31761']:  # TODO: удалить
        div_type, div_name = cell[5].value, cell[6].value
        if div_type in ['O', 'P', 'H', 'B']:
            div_full_name = f"{div_name} {obj_decode[div_type]}"
        else:
            div_full_name = f"{obj_decode[div_type]} {div_name}"

        atu_num = cell[obj_to_column[div_type]].value

        reg = int(atu_num[2:4])
        distr = int(atu_num[4:6])
        hrom = int(atu_num[6:9])
        munic = int(atu_num[9:12])
        distr_city = int(atu_num[12:14])
        div_num = int(atu_num[-5:])

        # If object of administrative division is region or city with special status
        if div_type in ['O', 'K']:
            curs.execute('INSERT INTO katotth(reg, div_num, div_type, div_name, div_full_name) '
                         'VALUES(%s, %s, %s, %s, %s)', (reg, div_num, obj_decode[div_type], div_name, div_full_name))
            # print(f"{reg} - {distr} - {hrom} - {munic} - {distr_city} : {div_num} - {div_full_name}")
        # If object of administrative division is district
        elif div_type == 'P':
            curs.execute('''INSERT INTO katotth(reg, distr, div_num, div_type, div_name, div_full_name) 
                                        VALUES(%s, %s, %s, %s, %s, %s)''',
                         (reg, distr, div_num, obj_decode[div_type], div_name, div_full_name))
            # print(f"{reg} - {distr} - {hrom} - {munic} - {distr_city} : {div_num} - {div_full_name}")
        # If object of administrative division is territorial hromada
        elif div_type == 'H':
            curs.execute('''INSERT INTO katotth(reg, distr, hrom, div_num, div_type, div_name, div_full_name) 
                                        VALUES(%s, %s, %s, %s, %s, %s, %s)''',
                         (reg, distr, hrom, div_num, obj_decode[div_type], div_name, div_full_name))
            # print(f"{reg} - {distr} - {hrom} - {munic} - {distr_city} : {div_num} - {div_full_name}")
        # If object of administrative division is one of the city or or village or special status
        elif div_type in ['M', 'T', 'C', 'X']:
            curs.execute('''INSERT INTO katotth(reg, distr, hrom, munic, div_num, div_type, div_name, div_full_name) 
                                        VALUES(%s, %s, %s, %s, %s, %s, %s, %s)''',
                         (reg, distr, hrom, munic, div_num, obj_decode[div_type], div_name, div_full_name))
            # print(f"{reg} - {distr} - {hrom} - {munic} - {distr_city} : {div_num} - {div_full_name}")
        # If object of administrative division is district in city
        elif div_type == 'B':
            curs.execute('''INSERT INTO katotth(reg, distr, hrom, munic, distr_city, div_num, 
                                                div_type, div_name, div_full_name) 
                            VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s)''',
                         (reg, distr, hrom, munic, distr_city, div_num, obj_decode[div_type], div_name, div_full_name))
            # print(f"{reg} - {distr} - {hrom} - {munic} - {distr_city} : {div_num} - {div_full_name}")

conn.commit()

workbook.close()

'''Columns: A - region(reg), B - district(dist), C - hromada(hrom), D - communities(comm), 
E - districts in cities(dist_city), F - object category(div_type), G - name'''

# [cell.value for cell in sheet['A']]  пробегает все значения в колонке А
# [cell[6].value for cell in sheet.rows]  пробегает по значениям ячеек конкретной колонки
