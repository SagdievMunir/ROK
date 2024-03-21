import tkinter
import sqlite3
import customtkinter
import tkinter.messagebox
import pandas as pd
import os
from datetime import datetime


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

conn = sqlite3.connect("database/ROK.db")

rootE = None
var = None

def EX1():
    rootE.destroy()
    E_display()


def EX2():
    rootB.destroy()
    SH_R()


def EX3():
    rootE.destroy()
    delo()


def EX4():
    rootDE.destroy()
    SH_R()


e1, e2, e3, e4 = None, None, None, None


def inp():
    global e1, e2, e3, e4, var
    e1 = t1.get()
    e2 = t2.get()
    e3 = t2_2.get()
    e4 = t3.get()
    conn.execute("INSERT INTO WORK VALUES(?,?,?,?)", (e1, e2, e3, e4))
    conn.commit()
    tkinter.messagebox.showinfo("ROKARM система", "Данные о штатном расписании сохранены")


# inp_s, errorS, t, i, q, dis1, dis2, dis3, dis4 = None, None, None, None, None, None, None, None, None
# dis5, l1, l2, l3, l4, l5 = None, None, None, None, None, None


def Search_button():
    global inp_s, entry, errorS, t, i, q, dis1, dis2, dis3, dis4, dis5
    global l1, l2, l3, l4, l5
    c1 = conn.cursor()
    inp_s = entry.get()
    p = list(c1.execute('SELECT * FROM WORK where DOCUMENT_ID=?', (inp_s,)))
    if len(p) == 0:
        if 'errorS' in globals():
            errorS.destroy()
        errorS = tkinter.Label(
            rootB,
            text="Данные о расписании не найдены",
            bg="#1A1919",
            fg='#daffda',
            font="Times 16 bold"
        )
        errorS.pack(fill='x')
    else:
        if 'errorS' in globals():
            errorS.destroy()  # Уничтожаем метку, если она существует
        t = c1.execute('SELECT * FROM WORK where DOCUMENT_ID=?', (inp_s,))
        # Удаление предыдущих меток
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
        if 'dis2' in globals():
            dis2.destroy()
        if 'dis3' in globals():
            dis3.destroy()
        if 'dis4' in globals():
            dis4.destroy()
        if 'dis5' in globals():
            dis5.destroy()

        for i in t:

            l2 = tkinter.Label(
                rootB,
                text="№ Документа",
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
                text="Код должности",
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
                text="Код отдела",
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
                text="Количество сотрудников",
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

            l2.pack(fill='x')
            dis2.pack(fill='x')
            l3.pack(fill='x')
            dis3.pack(fill='x')
            l4.pack(fill='x')
            dis4.pack(fill='x')
            l5.pack(fill='x')
            dis5.pack(fill='x')
            conn.commit()

def E_display():
    global rootB, head, inp_s, entry, searchB
    rootB = customtkinter.CTk()
    rootB.geometry('360x400+210+310')
    rootB.title("Окно поиска расписания")

    head = tkinter.Label(
        rootB,
        text="Введите № записи",
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


def get_next_app_id():
    c = conn.cursor()
    c.execute("SELECT MAX(DOCUMENT_ID) FROM WORK")
    max_id = c.fetchone()[0]
    return max_id + 1 if max_id else 1


# Exporting to Excel
def export_to_excel():
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f"work_{current_time}.xlsx"

    output_folder = "ExcelFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_path = os.path.join(output_folder, file_name)

    df = pd.read_sql_query("SELECT * FROM WORK", conn)
    df.to_excel(file_path, index=False)

    tkinter.messagebox.showinfo("ROKARM система", "Данные экспортированы в таблицу")


def SH_R():
    global rootE, t1, t2, t2_2, t3, lb, t4, t5, t6, var

    rootE = customtkinter.CTk()
    rootE.title("Добавление штатного расписания")
    rootE.geometry('360x350+220+315')
    rootE.resizable(width=False, height=False)
    rootE.configure(bg="#201E1F")

    var = tkinter.StringVar(master=rootE)

    next_id = get_next_app_id()

    H = tkinter.Label(
        rootE,
        text="Добавление расписния",
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
        text="№ записи"
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
        text="Код должности"
    )

    l2.place(x=10, y=110)

    t2 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t2.place(x=10, y=140)

    l3 = customtkinter.CTkLabel(
        rootE,
        text="Код отдела"
    )

    l3.place(x=10, y=170)

    t2_2 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t2_2.place(x=10, y=200)

    l4 = customtkinter.CTkLabel(
        rootE,
        text="Количество сотрудников"
    )

    l4.place(x=10, y=230)

    t3 = customtkinter.CTkEntry(
        rootE,
        width=180
    )

    t3.place(x=10, y=260)

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
    export_button.place(x=200, y=240)

    rootE.iconbitmap('assets/rok.ico')
    rootE.mainloop()


def delling():
    global d1, de
    de = str(d1.get())
    p = list(conn.execute("select * from WORK where DOCUMENT_ID=?", (de,)))
    if len(p) != 0:
        conn.execute("DELETE from WORK where DOCUMENT_ID=?", (de,))
        dme = tkinter.Label(
            rootDE,
            text="Расписание успешно удалено из базы",
            bg="#201E1F",
            fg="green"
        )

        dme.pack(fill='x')
        conn.commit()
    else:
        error = tkinter.Label(
            rootDE,
            text="Расписание не найдено",
            bg="#201E1F",
            fg="Red"
        )

        error.pack(fill='x')


rootDE = None


def delo():
    global rootDE, d1
    rootDE = customtkinter.CTk()
    rootDE.geometry("360x610+210+130")
    rootDE.title("Удаление расписания")
    rootDE.configure(bg="#201E1F")
    h1 = tkinter.Label(
        rootDE,
        text="Введите № записи расписания\nдля удаления из базы",
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
