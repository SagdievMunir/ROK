import tkinter
import sqlite3
import tkinter.messagebox
from PATDELSU import P_display
from PATDELSU import D_display
from PATDELSU import P_UPDATE
from RooMT import Room_all
from BILLING import BILLING
from TRIP import TRIP
from employee_reg import emp_screen
from appointment import SH_R
import customtkinter
from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd
import os

conn = sqlite3.connect("database/ROK.db")
print("Связь с БД установлена")

# variables
root1 = None
rootp = None
pat_ID = None
pat_name = None
pat_dob = None
pat_address = None
pat_sex = None
pat_BG = None
pat_email = None
pat_contact = None
pat_contactalt = None
pat_CT = None

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")


def export_to_excel():
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f"ENCOURAGEMENT_{current_time}.xlsx"

    output_folder = "ExcelFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_path = os.path.join(output_folder, file_name)

    df = pd.read_sql_query("SELECT * FROM ENCOURAGEMENT", conn)
    df.to_excel(file_path, index=False)

    tkinter.messagebox.showinfo("ROKARM система", "Данные экспортированы в таблицу")


def ex():
    global root1
    root1.destroy()


def export_all_tables_to_excel():
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f"all_tables{current_time}.xlsx"

    output_folder = "ExcelFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    excel_file = os.path.join(output_folder, file_name)

    # Создаем объект ExcelWriter для записи в один файл
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

    # Получаем список таблиц в базе данных
    tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()

    for table in tables:
        table_name = table[0]
        # Создаем объект DataFrame для каждой таблицы
        df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
        # Записываем DataFrame в свой лист Excel
        df.to_excel(writer, sheet_name=table_name, index=False)

    # Закрываем объект ExcelWriter
    writer.close()

    tkinter.messagebox.showinfo("ROKARM система", "Данные экспортированы в таблицу")


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry('{}x{}+{}+{}'.format(width, height, x, y))


def menu(center_window):
    global root1, button1, button2, button3, button4, button5, m, button6, button7
    root1 = customtkinter.CTk()
    root1.geometry("750х470")
    root1.title("Главное меню")
    root1.resizable(width=False, height=False)
    root1.configure()

    frame1 = customtkinter.CTkFrame(root1, fg_color="#201E1F")
    frame1.pack(expand=True, fill="both")

    center_window(root1, 750, 470)

    m = tkinter.Label(
        frame1,
        text="Главное меню",
        font=("Times", 16, "bold"),
        fg="#daffda",
        bg="#1A1919",
        padx=20,
        pady=10
    )

    button1 = customtkinter.CTkButton(
        root1,
        text="1. Сотрудники",
        font=("Times", 18, "bold"),
        width=700,
        command=emp_screen
    )

    button2 = customtkinter.CTkButton(
        root1,
        text="2. Поощрения",
        font=("Times", 18, "bold"),
        width=700,
        command=PAT
    )

    button3 = customtkinter.CTkButton(
        root1,
        text="3. Отпуска",
        font=("Times", 18, "bold"),
        width=700,
        command=Room_all
    )

    button4 = customtkinter.CTkButton(
        root1,
        text="4. Штатное расписание",
        font=("Times", 18, "bold"),
        width=700,
        command=SH_R
    )

    button5 = customtkinter.CTkButton(
        root1,
        text="5. Повышение квалификации",
        font=("Times", 18, "bold"),
        width=700,
        command=BILLING
    )

    button6 = customtkinter.CTkButton(
        root1,
        text="6. Командировки",
        font=("Times", 18, "bold"),
        width=700,
        command=TRIP
    )

    button7 = customtkinter.CTkButton(
        root1,
        text="7. Экспорт всех таблиц в EXCEL",
        font=("Times", 18, "bold"),
        width=700,
        command=export_all_tables_to_excel
    )

    button8 = customtkinter.CTkButton(
        root1,
        text="Выйти",
        font=("Times", 18, "bold"),
        width=700,
        command=ex
    )

    m.pack(fill='x')
    button1.pack(side=tkinter.TOP)
    button1.place(x=20, y=70)
    button2.pack(side=tkinter.TOP)
    button2.place(x=20, y=120)
    button3.pack(side=tkinter.TOP)
    button3.place(x=20, y=170)
    button4.pack(side=tkinter.TOP)
    button4.place(x=20, y=220)
    button5.pack(side=tkinter.TOP)
    button5.place(x=20, y=270)
    button6.pack(side=tkinter.TOP)
    button6.place(x=20, y=320)
    button7.pack(side=tkinter.TOP)
    button7.place(x=20, y=370)
    button8.pack(side=tkinter.TOP)
    button8.place(x=20, y=420)

    root1.iconbitmap('assets/rok.ico')
    root1.mainloop()


p = None


def IN_PAT():
    global pp1, pp2, pp3, pp4, pp5, pp6, pp7, pp8, pp9, pp10, ce1, inp_s, conn
    global pat_address, pat_BG, pat_contact, pat_contactalt, pat_CT, pat_dob, pat_email, pat_ID, pat_name, pat_sex
    conn = sqlite3.connect("database/ROK.db")
    conn.cursor()
    pp1 = pat_ID.get()
    pp2 = pat_name.get()
    pp3 = pat_sex.get()
    pp4 = pat_BG.get()
    pp5 = pat_dob.get()
    pp7 = pat_contactalt.get()
    pp8 = pat_email.get()
    conn.execute('INSERT INTO ENCOURAGEMENT (ENPLOYEE_ID, DOCUMENT_ID, MOTIVE, TYPE_ENCOURAGEMENT, DATE, EXPLANATION, '
                 'AMOUNT) VALUES(?,?,?,?,?,?,?)', (pp1, pp2, pp3, pp4, pp5, pp7, pp8))

    c1 = conn.cursor()
    inp_s = pat_name.get()
    t = c1.execute('SELECT * FROM EPLOYEE where ENPLOYEE_ID=?', (inp_s,))

    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_date = datetime.now().strftime("%d.%m.%Y")

    context = {'doc_number': pp1,
               'doc_date': file_date,
               's_id': pp2,
               'explanation': pp3,
               'type_encouragement': pp4,
               'amount': pp8}
    for i in t:
        context['fullname'] = i[1]
        context['post'] = i[4]
    doc = DocxTemplate("PremiaShablon.docx")
    doc.render(context)

    output_folder = "PremiaFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_name = f"generated_{current_time}.docx"
    file_path = os.path.join(output_folder, file_name)
    doc.save(file_path)

    tkinter.messagebox.showinfo("ROKARM система", "Данные добавлены в систему")
    conn.commit()


def EXO():
    rootp.destroy()


def EX1():
    rootp.destroy()
    P_display()


def EX2():
    rootp.destroy()
    D_display()


def EX3():
    rootp.destroy()
    P_UPDATE()


def nothing():
    print("Связь с базой данных установлена ")


def nothing1():
    print("ROK ARM")


back = None
SEARCH = None
DELETE = None
UPDATE = None


def get_next_encouragement_id():
    c = conn.cursor()
    c.execute("SELECT MAX(DOCUMENT_ID) FROM ENCOURAGEMENT")
    max_id = c.fetchone()[0]
    return max_id + 1 if max_id else 1


def PAT():
    global pat_address, pat_BG, pat_contact, pat_contactalt, pat_CT, pat_dob, pat_email, pat_ID, pat_name, pat_sex
    global rootp, regform, id_1, name, dob, sex, email, ct, addr, c1, c2, bg, SUBMIT, menubar, filemenu, back, SEARCH, DELETE, UPDATE
    rootp = customtkinter.CTk()
    rootp.title("Поощрения")
    rootp.geometry('360x530+1350+130')
    rootp.resizable(width=False, height=False)
    menubar = tkinter.Menu(rootp)

    next_id = get_next_encouragement_id()

    filemenu = tkinter.Menu(
        menubar,
        tearoff=0
    )

    filemenu.add_command(
        label="Добавить",
        command=PAT
    )

    filemenu.add_separator()

    filemenu.add_command(
        label="Выйти",
        command=EXO
    )

    helpmenu = tkinter.Menu(
        menubar,
        tearoff=0
    )

    helpmenu.add_command(
        label="Помощь",
        command=nothing
    )

    helpmenu.add_command(
        label="Подробнее",
        command=nothing1
    )

    menubar.add_cascade(
        label="Файл",
        menu=filemenu
    )

    menubar.add_cascade(
        label="Помощь",
        menu=helpmenu
    )

    rootp.config(
        menu=menubar
    )

    regform = tkinter.Label(
        rootp,
        text="Поощрения",
        bg="#1A1919",
        fg='#daffda',
        font="Times 16 bold",
        padx=20,
        pady=10
    )

    regform.place(x=0, y=0)
    regform.pack(fill='x')

    id_1 = customtkinter.CTkLabel(
        rootp,
        width=180,
        text="ID Записи"
    )

    id_1.place(x=10, y=50)

    pat_ID = customtkinter.CTkEntry(
        rootp,
        width=180,
    )
    pat_ID.insert(0, next_id)
    pat_ID.place(x=10, y=80)

    name = customtkinter.CTkLabel(
        rootp,
        width=180,
        text="ID Сотрудника"
    )

    name.place(x=10, y=110)

    pat_name = customtkinter.CTkEntry(
        rootp,
        width=180,
    )

    pat_name.place(x=10, y=140)

    sex = customtkinter.CTkLabel(
        rootp,
        width=180,
        text="Мотив поощрения"
    )

    sex.place(x=10, y=170)

    pat_sex = customtkinter.CTkEntry(
        rootp,
        width=180,
    )

    pat_sex.place(x=10, y=200)

    dob = customtkinter.CTkLabel(
        rootp,
        width=180,
        text="Дата (ГГГГ-ММ-ДД)"
    )

    dob.place(x=10, y=230)

    pat_dob = customtkinter.CTkEntry(
        rootp,
        width=180,
    )

    pat_dob.place(x=10, y=260)

    bg = customtkinter.CTkLabel(
        rootp,
        width=180,
        text="Вид поощрения"
    )

    bg.place(x=10, y=290)

    pat_BG = customtkinter.CTkEntry(
        rootp,
        width=180,
    )

    pat_BG.place(x=10, y=320)

    c2 = customtkinter.CTkLabel(
        rootp,
        width=180,
        text="Основание"
    )

    c2.place(x=10, y=350)

    pat_contactalt = customtkinter.CTkEntry(
        rootp,
        width=180,
    )

    pat_contactalt.place(x=10, y=380)

    email = customtkinter.CTkLabel(
        rootp,
        width=180,
        text="Сумма"
    )

    email.place(x=10, y=410)

    pat_email = customtkinter.CTkEntry(
        rootp,
        width=180,
    )

    pat_email.place(x=10, y=440)

    back = customtkinter.CTkButton(
        rootp,
        text="Назад",
        command=EXO
    )

    back.place(x=200, y=80)

    SEARCH = customtkinter.CTkButton(
        rootp,
        text="Поиск",
        command=EX1
    )

    SEARCH.place(x=200, y=120)

    DELETE = customtkinter.CTkButton(
        rootp,
        text="Удалить",
        command=EX2
    )

    DELETE.place(x=200, y=160)

    UPDATE = customtkinter.CTkButton(
        rootp,
        text="Обновить",
        command=EX3
    )

    UPDATE.place(x=200, y=200)

    export_button = customtkinter.CTkButton(
        rootp,
        text="Экспорт в Excel",
        command=export_to_excel
    )
    export_button.place(x=200, y=240)

    SUBMIT = customtkinter.CTkButton(
        rootp,
        width=180,
        text="Подтвердить",
        command=IN_PAT
    )

    SUBMIT.place(x=10, y=490)

    regform.pack(fill='x')
    rootp.iconbitmap('assets/rok.ico')

    rootp.mainloop()
