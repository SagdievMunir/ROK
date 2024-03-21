import tkinter
import sqlite3
import customtkinter
import tkinter.messagebox
from docxtpl import DocxTemplate
from datetime import datetime
import os
import pandas as pd
import matplotlib.pyplot as plt

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

conn = sqlite3.connect("database/ROK.db")

rootE = None
var = None


def export_to_excel():
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f"EPLOYEE_{current_time}.xlsx"

    output_folder = "ExcelFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_path = os.path.join(output_folder, file_name)

    df = pd.read_sql_query("SELECT * FROM EPLOYEE", conn)
    df.to_excel(file_path, index=False)

    tkinter.messagebox.showinfo("ROKARM система", "Данные экспортированы в таблицу")


def plot_employee_experience():

    c = conn.cursor()

    c.execute("SELECT WORK_EXPERIENCE FROM EPLOYEE")
    rows = c.fetchall()

    experience_list = [row[0] for row in rows]

    experience_counts = {}
    for experience in experience_list:
        if experience in experience_counts:
            experience_counts[experience] += 1
        else:
            experience_counts[experience] = 1

    experiences = sorted(experience_counts.keys())
    counts = [experience_counts[exp] for exp in experiences]

    fig, ax = plt.subplots()
    fig.patch.set_facecolor('#272727')

    ax.set_facecolor('#272727')
    plt.rcParams['figure.facecolor'] = '#272727'
    plt.plot(experiences, counts, marker='o', color='#2fa572', linestyle='-')
    plt.xlabel('Стаж работы (года)', color='white')
    plt.ylabel('Количество сотрудников', color='white')
    plt.title('Распределение стажа работы сотрудников', color='white')
    plt.grid(linewidth=0.2)
    plt.xlim(0,)
    plt.ylim(0,)
    plt.tick_params(axis='x', colors='white')
    plt.tick_params(axis='y', colors='white')

    plt.locator_params(axis='y', integer=True)
    plt.yticks(list(range(int(min(counts)), int(max(counts)) + 1, 1)))

    plt.tight_layout()

    plt.show()


def plot_employee_positions():
    c = conn.cursor()

    c.execute("SELECT PHONE_NUMBER FROM EPLOYEE")
    rows = c.fetchall()

    position_counts = {}
    for row in rows:
        position = row[0]
        if position in position_counts:
            position_counts[position] += 1
        else:
            position_counts[position] = 1

    positions = list(position_counts.keys())
    counts = list(position_counts.values())

    fig, ax = plt.subplots()
    fig.patch.set_facecolor('#272727')

    ax.set_facecolor('#272727')
    plt.rcParams['figure.facecolor'] = '#272727'

    plt.bar(positions, counts, color='#2fa572')
    plt.xlabel('Должность', color='white')
    plt.ylabel('Количество сотрудников', color='white')
    plt.title('Количество сотрудников по должностям', color='white')
    plt.xticks(rotation=45, ha='right')
    plt.tick_params(axis='x', colors='white')
    plt.tick_params(axis='y', colors='white')
    plt.grid(linewidth=0.2)
    plt.yticks(list(range(int(min(counts)), int(max(counts)) + 1, 1)))
    plt.tight_layout()

    plt.show()


def EX1():
    rootE.destroy()
    E_display()


def EX2():
    rootB.destroy()
    emp_screen()


def EX3():
    rootE.destroy()
    delo()


def EX4():
    rootDE.destroy()
    emp_screen()


def inp():
    global e1, e2, e3, e4, e5, e6, e7, e8, e9, var
    e1 = t1.get()
    e2 = t2.get()
    e3 = t2_2.get()
    e4 = t3.get()
    e5 = lb.get()
    e6 = t4.get()
    e7 = t5.get()
    e8 = t6.get()
    e9 = t7.get()
    conn.execute("INSERT INTO EPLOYEE VALUES(?,?,?,?,?,?,?,?,?)", (e1, e2, e3, e4, e5, e6, e7, e8, e9,))

    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_date = datetime.now().strftime("%d.%m.%Y")

    context = {'fullname': e2,
               'post': e5,
               'phone': e6,
               'date': file_date}
    doc = DocxTemplate("PriemkaShablon.docx")
    doc.render(context)

    output_folder = "PriemkaFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_name = f"generated_{current_time}.docx"
    file_path = os.path.join(output_folder, file_name)
    doc.save(file_path)

    conn.commit()
    tkinter.messagebox.showinfo("ROKARM система", "Данные о работнике занесены")


def Search_button():
    global inp_s, entry, errorS, t, i, q, dis1, dis2, dis3, dis4, dis5, dis6, dis7, dis8, dis9, dis10
    global l1, l2, l3, l4, l5, l6, l7, l8, l9, l10
    c1 = conn.cursor()
    inp_s = entry.get()

    p = list(c1.execute('select * from EPLOYEE where FULLNAME=?', (inp_s,)))

    if len(p) == 0:
        if 'errorS' in globals():
            errorS.destroy()
        errorS = tkinter.Label(
            rootB,
            text="Данные о записи не найдены",
            bg="#1A1919",
            fg='#daffda',
            font="Times 16 bold"
        )
        errorS.pack(fill='x')
    else:
        if 'errorS' in globals():
            errorS.destroy()
        t = c1.execute('SELECT * FROM EPLOYEE where FULLNAME=?', (inp_s,))

        if 'errorS' in globals():
            errorS.destroy()
        if 'l2' in globals():
            l2.destroy()
        if 'l3' in globals():
            l3.destroy()
        if 'l4' in globals():
            l4.destroy()
        if 'l5' in globals():
            l5.destroy()
        if 'l6' in globals():
            l6.destroy()
        if 'l7' in globals():
            l7.destroy()
        if 'l8' in globals():
            l8.destroy()
        if 'l9' in globals():
            l9.destroy()
        if 'l10' in globals():
            l10.destroy()
        if 'dis2' in globals():
            dis2.destroy()
        if 'dis3' in globals():
            dis3.destroy()
        if 'dis4' in globals():
            dis4.destroy()
        if 'dis5' in globals():
            dis5.destroy()
        if 'dis6' in globals():
            dis6.destroy()
        if 'dis7' in globals():
            dis7.destroy()
        if 'dis8' in globals():
            dis8.destroy()
        if 'dis9' in globals():
            dis9.destroy()
        if 'dis10' in globals():
            dis10.destroy()

        for i in t:

            l2 = tkinter.Label(
                rootB,
                text="ID Сотрудника",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis2 = tkinter.Label(
                rootB,
                text=i[0],
                fg='white',
                bg="#1A1919",
            )

            l3 = tkinter.Label(
                rootB,
                text="ФИО сотрудника",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis3 = tkinter.Label(
                rootB,
                fg='white',
                bg="#1A1919",
                text=i[1]
            )

            l4 = tkinter.Label(
                rootB,
                text="ИНН",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis4 = tkinter.Label(
                rootB,
                fg='white',
                bg="#1A1919",
                text=i[2],
            )

            l5 = tkinter.Label(
                rootB,
                text="СНИЛС",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis5 = tkinter.Label(
                rootB,
                fg='white',
                bg="#1A1919",
                text=i[3]
            )

            l6 = tkinter.Label(
                rootB,
                text="Должность",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis6 = tkinter.Label(
                rootB,
                fg='white',
                bg="#1A1919",
                text=i[4]
            )

            l7 = tkinter.Label(
                rootB,
                text="Номер телефона",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis7 = tkinter.Label(
                rootB,
                fg='white',
                bg="#1A1919",
                text=i[5]
            )

            l8 = tkinter.Label(
                rootB,
                text="День рождения",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis8 = tkinter.Label(
                rootB,
                fg='white',
                bg="#1A1919",
                text=i[6]
            )

            l9 = tkinter.Label(
                rootB,
                text="Стаж работы",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis9 = tkinter.Label(
                rootB,
                fg='white',
                bg="#1A1919",
                text=i[7]
            )

            l10 = tkinter.Label(
                rootB,
                text="Зарплата",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis10 = tkinter.Label(
                rootB,
                fg='white',
                bg="#1A1919",
                text=i[8]
            )

            l2.pack(fill='x')
            dis2.pack(fill='x')
            l3.pack(fill='x')
            dis3.pack(fill='x')
            l4.pack(fill='x')
            dis4.pack(fill='x')
            l5.pack(fill='x')
            dis5.pack(fill='x')
            l6.pack(fill='x')
            dis6.pack(fill='x')
            l7.pack(fill='x')
            dis7.pack(fill='x')
            l8.pack(fill='x')
            dis8.pack(fill='x')
            l9.pack(fill='x')
            dis9.pack(fill='x')
            l10.pack(fill='x')
            dis10.pack(fill='x')
            conn.commit()

def E_display():
    global rootB, head, inp_s, entry, searchB
    rootB = customtkinter.CTk()
    rootB.geometry('360x610+210+130')
    rootB.title("Окно поиска")

    head = tkinter.Label(
        rootB,
        text="Введите ФИО сотрудника",
        bg="#1A1919",
        fg='#daffda',
        font=("Times", 16, "bold"),
        padx=20,
        pady=10
    )

    entry = customtkinter.CTkEntry(
        rootB,
        width=180,
    )

    searchB = customtkinter.CTkButton(
        rootB,
        text='Поиск',
        width=180,
        command=Search_button
    )

    backB = customtkinter.CTkButton(
        rootB,
        text='Назад',
        width=180,
        command=EX2
    )

    head.pack(fill='x')
    entry.pack(pady=5)
    searchB.pack(pady=5)
    backB.pack(pady=5)
    rootB.iconbitmap('assets/rok.ico')
    rootB.mainloop()

def ex():
    rootE.destroy()


def get_next_employee_id():
    c = conn.cursor()
    c.execute("SELECT MAX(ENPLOYEE_ID) FROM EPLOYEE")
    max_id = c.fetchone()[0]
    return max_id + 1 if max_id else 1


def emp_screen():
    global rootE, t1, t2, t2_2, t3, lb, t4, t5, t6, t7, var

    rootE = customtkinter.CTk()
    rootE.title("Регистрация сотрудника")
    rootE.geometry('360x610+210+130')
    rootE.resizable(width=False, height=False)
    rootE.configure(bg="#201E1F")

    var = tkinter.StringVar(master=rootE)

    next_id = get_next_employee_id()

    H = tkinter.Label(
        rootE,
        text="Регистрация сотрудника",
        bg="#1A1919",
        fg='#daffda',
        font=("Times", 16, "bold"),
        padx=20,
        pady=10
    )

    H.place(x=0, y=0)
    H.pack(fill='x')

    l1 = customtkinter.CTkLabel(
        rootE,
        text="ID сотрудника"
    )

    l1.place(x=10, y=50)

    t1 = customtkinter.CTkEntry(
        rootE,
        width=180
    )
    t1.insert(0, next_id)
    t1.place(x=10, y=80)

    l2 = customtkinter.CTkLabel(
        rootE,
        text="ФИО сотрудника"
    )

    l2.place(x=10, y=110)

    t2 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t2.place(x=10, y=140)

    l3 = customtkinter.CTkLabel(
        rootE,
        text="ИНН"
    )

    l3.place(x=10, y=170)

    t2_2 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t2_2.place(x=10, y=200)

    l4 = customtkinter.CTkLabel(
        rootE,
        text="СНИЛС"
    )

    l4.place(x=10, y=230)

    t3 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t3.place(x=10, y=260)

    l5 = customtkinter.CTkLabel(
        rootE,
        text="Должность"
    )

    l5.place(x=10, y=290)

    lb = customtkinter.CTkComboBox(
        rootE,
        width=180,
        values=["Разработчик ПО",
                "Веб-разработчик",
                "Аналитик",
                "Сетевой администратор",
                "Администратор БД",
                "Облачный инженер",
                "Специалист по кибербезу"]
    )

    lb.place(x=10, y=320)

    l6 = customtkinter.CTkLabel(
        rootE,
        text="Номер телефона"
    )

    l6.place(x=10, y=350)

    t4 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t4.place(x=10, y=380)

    l7 = customtkinter.CTkLabel(
        rootE,
        text="День рождения (ГГГГ-ММ-ДД)"
    )

    l7.place(x=10, y=410)

    t5 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t5.place(x=10, y=440)

    l8 = customtkinter.CTkLabel(
        rootE,
        text="Стаж работы"
    )

    l8.place(x=10, y=470)

    t6 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t6.place(x=10, y=500)

    l9 = customtkinter.CTkLabel(
        rootE,
        text="Зарплата"
    )

    l9.place(x=10, y=530)

    t7 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t7.place(x=10, y=560)

    b1 = customtkinter.CTkButton(
        rootE,
        text="Сохранить",
        command=inp
    )

    b1.place(x=200, y=80)

    b2 = customtkinter.CTkButton(
        rootE,
        text="Удалить",
        command=EX3
    )

    b2.place(x=200, y=120)

    b3 = customtkinter.CTkButton(
        rootE,
        text="Выйти",
        command=ex
    )

    b3.place(x=200, y=200)

    SEARCH_N = customtkinter.CTkButton(
        rootE,
        text="Поиск",
        command=EX1
    )

    SEARCH_N.place(x=200, y=160)

    export_button = customtkinter.CTkButton(
        rootE,
        text="Экспорт в Excel",
        command=export_to_excel
    )
    export_button.place(x=200, y=455)

    plot_sec = customtkinter.CTkButton(
        rootE,
        text="График сотрудников\n по стажу",
        command=plot_employee_experience
    )
    plot_sec.place(x=200, y=505)

    plot_button = customtkinter.CTkButton(
        rootE,
        text="График сотрудников\n по должностям",
        command=plot_employee_positions
    )
    plot_button.place(x=200, y=555)

    rootE.iconbitmap('assets/rok.ico')
    rootE.mainloop()


def delling():
    global d1, de
    de = str(d1.get())
    p = list(conn.execute("select * from EPLOYEE where ENPLOYEE_ID=?", (de,)))
    if len(p) != 0:
        conn.execute("DELETE from EPLOYEE where ENPLOYEE_ID=?", (de,))
        dme = tkinter.Label(
            rootDE,
            text="Сотрудник успешно удален из базы",
            bg="#201E1F",
            fg="green"
        )

        dme.pack(fill='x')
        conn.commit()
    else:
        error = tkinter.Label(
            rootDE,
            text="Сотрудник не найден",
            bg="#201E1F",
            fg="Red"
        )

        error.pack(fill='x')


rootDE = None


def delo():
    global rootDE, d1
    rootDE = customtkinter.CTk()
    rootDE.geometry("360x610+210+130")
    rootDE.title("Удаление сотрудника")
    rootDE.configure(bg="#201E1F")
    h1 = tkinter.Label(
        rootDE,
        text="Введите ID работника\nдля удаления из базы",
        bg="#1A1919",
        fg='#daffda',
        font="Times 16 bold",
        padx=20,
        pady=10
    )

    d1 = customtkinter.CTkEntry(
        rootDE,
        width=180,
    )

    B1 = customtkinter.CTkButton(
        rootDE,
        text="Удалить",
        command=delling,
        width=180,
    )

    B2 = customtkinter.CTkButton(
        rootDE,
        text="Назад",
        command=EX4,
        width=180,
    )

    h1.pack(fill='x')
    d1.pack(pady=5)
    B1.pack(pady=5)
    B2.pack(pady=5)
    rootDE.iconbitmap('assets/rok.ico')
    rootDE.mainloop()
