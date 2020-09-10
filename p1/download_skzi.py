from openpyxl import load_workbook
import psycopg2

orderSkziFile = 'C:\\microservices\\skzi_base.xlsx'

skzi = []

db_config = {'database':'OKZ_DB',
           'user':'root',
           'password':'test',
           'host':'127.0.0.1'}

conn = psycopg2.connect(**db_config)
cursor = conn.cursor()

try:
    cursor.execute('''CREATE TABLE Skzis
    (skziuser_name CHAR (512),
    skziserial_name CHAR (255),
    pvmserial_nam CHAR (255),
    pechat_number CHAR (255),
    adress CHAR (1024),
    szz CHAR (255),
    catalognumber CHAR (255),
    description CHAR (1024),
    sertificate CHAR (255),
    skzi_type CHAR (1024),
    ais CHAR (1024),
    software CHAR (1024),
    organization CHAR (1024));''')

except psycopg2.errors.DuplicateTable:
    print ('Таблица в базе данных существует')

conn.commit()
conn.close()

'''
class Skzis:

    def __init__(self, skziuser_name, 
                skziserial_name,
                pvmserial_nam,
                pechat_number,
                adress,
                szz,
                catalognumber,
                description,
                sertificate,
                skzi_type,
                ais,
                software,
                organization,
                ):
        
        self.skziuser_name = skziuser_name
        self.skziserial_name = skziserial_name
        self.pvmserial_nam = pvmserial_nam
        self.pechat_number = pechat_numbe
        self.adress = adress
        self.szz = szz
        self.catalognumber = catalognumber
        self.description = description
        self.sertificate = sertificate
        self.skzi_type = skzi_type
        self.ais = ais
        self.software = software
        self.organization = organization
'''
def readSkzis():
    wb = load_workbook(orderSkziFile, data_only=True)
    #ws = wb['Сх криптозащСКЗИ']

    conn = psycopg2.connect(**db_config)
    cursor = conn.cursor()


    for row_num in range (2, ws.max_row):

        if (ws.cell(row = row_num, column = 3).value) == None or (ws.cell(row = row_num, column = 3).value).strip() =='':
            skziuser_name = 'Null'
        else:
            skziuser_name = (ws.cell(row = row_num, column = 3).value).strip()

        if (ws.cell(row = row_num, column = 5).value) == None or (ws.cell(row = row_num, column = 5).value).strip() =='':
            skziserial_name = 'Null'
        else:
            skziserial_name = (ws.cell(row = row_num, column = 5).value).strip()

        if (ws.cell(row = row_num, column = 7).value) == None or (ws.cell(row = row_num, column = 7).value).strip() =='':
            pvmserial_nam = 'Null'
        else:
            pvmserial_nam = (ws.cell(row = row_num, column = 7).value).strip()

        if (ws.cell(row = row_num, column = 8).value) == None or (ws.cell(row = row_num, column = 8).value).strip() =='':
            pechat_number = 'Null'
        else:
            pechat_number = (ws.cell(row = row_num, column = 8).value).strip()

        if (ws.cell(row = row_num, column = 9).value) == None or (ws.cell(row = row_num, column = 9).value).strip() =='':
            adress = 'Null'
        else:
            adress = (ws.cell(row = row_num, column = 9).value).strip()

        if (ws.cell(row = row_num, column = 11).value) == None or (ws.cell(row = row_num, column = 11).value).strip() =='':
            szz = 'Null'
        else:
            szz = (ws.cell(row = row_num, column = 11).value).strip()

        if (ws.cell(row = row_num, column = 12).value) == None or ((str(ws.cell(row = row_num, column = 12).value))).strip() =='':
            catalognumber = 'Null'
        else:
            catalognumber = (str(ws.cell(row = row_num, column = 12).value)).strip()

        if (ws.cell(row = row_num, column = 13).value) == None or (ws.cell(row = row_num, column = 13).value).strip() =='':
            description = 'Null'
        else:
            description = (ws.cell(row = row_num, column = 13).value).strip()

        if (ws.cell(row = row_num, column = 14).value) == None or (ws.cell(row = row_num, column = 14).value).strip() =='':
            sertificate = 'Null'
        else:
            sertificate = (ws.cell(row = row_num, column = 14).value).strip()

        if (ws.cell(row = row_num, column = 4).value) == None or (ws.cell(row = row_num, column = 4).value).strip() =='':
            skzi_type = 'Null'
        else:
            skzi_type = (ws.cell(row = row_num, column = 4).value).strip()

        if (ws.cell(row = row_num, column = 6).value) == None or (ws.cell(row = row_num, column = 6).value).strip() =='':
            ais = 'Null'
        else:
            ais = (ws.cell(row = row_num, column = 6).value).strip()

        if (ws.cell(row = row_num, column = 10).value) == None or (ws.cell(row = row_num, column = 10).value).strip() =='':
            software = 'Null'
        else:
            software = (ws.cell(row = row_num, column = 10).value).strip()

        if (ws.cell(row = row_num, column = 2).value) == None or (ws.cell(row = row_num, column = 2).value).strip() =='':
            organization = 'Null'
        else:
            organization = (ws.cell(row = row_num, column = 2).value).strip()

        
        #newSkzis = Skzis(skziuser_name, skziserial_name, pvmserial_nam, pechat_number, adress, szz, catalognumber, description, sertificate, skzi_type, ais, software, organization)
        #skzi.append(newSkzis)
        

        _SQL = """INSERT INTO Skzis (skziuser_name, skziserial_name, pvmserial_nam, pechat_number, adress, szz, catalognumber, description, sertificate, skzi_type, ais, software, organization) VALUES
        ('%(skziuser_name)s', '%(skziserial_name)s', '%(pvmserial_nam)s', '%(pechat_number)s', '%(adress)s', '%(szz)s', '%(catalognumber)s', '%(description)s', '%(sertificate)s', '%(skzi_type)s', '%(ais)s', '%(software)s', '%(organization)s')
        """%{'skziuser_name':skziuser_name, 'skziserial_name':skziserial_name, 'pvmserial_nam':pvmserial_nam, 'pechat_number':pechat_number, 'adress':adress, 'szz':szz, 'catalognumber':catalognumber, 'description':description, 'sertificate':sertificate, 'skzi_type':skzi_type, 'ais':ais, 'software':software, 'organization':organization}


        cursor.execute(_SQL)
        conn.commit()
    conn.close()

readSkzis()