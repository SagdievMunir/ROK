import sqlite3

conn = sqlite3.connect('ROK.db')
print("Связь с БД установлена")

conn.execute("Drop table if EXISTS EPLOYEE")
conn.execute("Drop table if EXISTS TRAINING")
conn.execute("Drop table if EXISTS WORK")
conn.execute("DROP TABLE if EXISTS ENCOURAGEMENT")
conn.execute("DROP TABLE if EXISTS TRIP")
conn.execute("DROP TABLE if EXISTS VACATION")


conn.execute("""CREATE TABLE EPLOYEE
                (ENPLOYEE_ID int(10) primary key,
                FULLNAME VARCHAR(255) not null,
                INN varchar(12) not null,
                SNILS varchar(11) not null,
                PHONE_NUMBER varchar(12) not null,
                DOB date not null,
                POST varchar(255) not null,
                WORK_EXPERIENCE int(3) not null,
                WAGES int(7) not null
             )""")
print("Таблица EPLOYEE успешно создана")


conn.execute("""CREATE TABLE TRAINING
                 (ENPLOYEE_ID int(10) not NULL,
                 DOCUMENT_ID int(20) PRIMARY KEY,
                 TRAINING varchar(255) not null,
                 DATE date not null,
                 TYPE varchar(20) not null,
                 INSTITUTION_NAME varchar(255) not null,
                 FOREIGN KEY(ENPLOYEE_ID) REFERENCES EPLOYEE(ENPLOYEE_ID));
            """)
print("Таблица TRAINING успешно создана")


conn.execute("""CREATE TABLE WORK
                 (DOCUMENT_ID int(10) primary key,
                 POST_CODE int(10) not NULL,
                 DEPARTAMENT_ID int(10) not NULL,
                 NUMBER_OF_EMPLOYEES int(10) not NULL
             )""")
print("Таблица TABLE успешно создана")


conn.execute("""CREATE TABLE ENCOURAGEMENT
                (ENPLOYEE_ID int(10) not NULL,
                 DOCUMENT_ID int(10) primary key,
                 MOTIVE varchar(20) not NULL,
                 TYPE_ENCOURAGEMENT varchar(20) not NULL,
                 DATE date not null,
                 EXPLANATION varchar(255) not null,
                 AMOUNT int(7) not null,
                 FOREIGN KEY(ENPLOYEE_ID) REFERENCES EPLOYEE(ENPLOYEE_ID));
             """)
print("Таблица ENCOURAGEMENT успешно создана")


conn.execute("""CREATE TABLE TRIP
                (ENPLOYEE_ID int(10) primary key,
                 DOCUMENT_ID int(10) not NULL,
                 COUNTRY_CITY varchar(50) not NULL,
                 ORDER_NUMBER int(10) not NULL,
                 DATE date not null,
                 FROM_DATE date not null,
                 BY_DATE date not null,
                 PURPOSE varchar(255) not null,
                 FOREIGN KEY(ENPLOYEE_ID) REFERENCES EPLOYEE(ENPLOYEE_ID));
             """)
print("Таблица TRIP успешно создана")


conn.execute("""CREATE TABLE VACATION
                (ENPLOYEE_ID int(10) not NULL,
                 DOCUMENT_ID int(10) primary key,
                 VACATION_TYPE varchar(20) not NULL,
                 DATE date not null,
                 NUMBER_OF_DAYS int(10) not NULL,
                 EXPLANATION varchar(255) not null,
                 FOREIGN KEY(ENPLOYEE_ID) REFERENCES EPLOYEE(ENPLOYEE_ID));
             """)
print("Таблица VACATION успешно создана")