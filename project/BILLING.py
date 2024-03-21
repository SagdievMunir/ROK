import sqlite3
import tkinter
import tkinter.messagebox
import customtkinter
from CTkListbox import *
from docxtpl import DocxTemplate
from datetime import datetime
import os
import matplotlib.pyplot as plt
import pandas as pd


conn = sqlite3.connect("database/ROK.db")

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

rootB = None
c1 = conn.cursor()


def export_to_excel():
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f"TRAINING_{current_time}.xlsx"

    output_folder = "ExcelFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_path = os.path.join(output_folder, file_name)

    df = pd.read_sql_query("SELECT * FROM TRAINING", conn)
    df.to_excel(file_path, index=False)

    tkinter.messagebox.showinfo("ROKARM система", "Данные экспортированы в таблицу")


def plot_training_positions():
    c = conn.cursor()

    c.execute("SELECT TYPE FROM TRAINING")
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
    plt.xlabel('Организация', color='white')
    plt.ylabel('Количество сотрудников', color='white')
    plt.title('Повышение квалификации по организациям', color='white')
    plt.xticks(rotation=45, ha='right')
    plt.grid(linewidth=0.2)
    plt.tick_params(axis='x', colors='white')
    plt.tick_params(axis='y', colors='white')
    plt.yticks(list(range(int(min(counts)), int(max(counts)) + 1, 1)))
    plt.tight_layout()

    plt.show()


def Search_button():
    global inp_s, entry, errorS, t, i, q, dis1, dis2, dis3, dis4, dis5, dis6, dis7, dis8, dis9, dis10
    global l1, l2, l3, l4, l5, l6, l7, l8, l9, l10
    c1 = conn.cursor()
    inp_s = entry.get()
    p = list(c1.execute('select * from TRAINING where DOCUMENT_ID=?', (inp_s,)))
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
        t = c1.execute('SELECT * FROM TRAINING where DOCUMENT_ID=?', (inp_s,))
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
                text="№ Записи",
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
                text="Дополнительные сведения",
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
                text="Дата (ГГГГ-ММ-ДД)",
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
                text="Наименование учреждения",
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
                text="Вид",
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
            conn.commit()

def B_display():
    global rootB, head, inp_s, entry, searchB
    rootB = customtkinter.CTk()
    rootB.geometry('520x550+60+190')
    rootB.title("Окно поиска")

    head = tkinter.Label(
        rootB,
        text="Введите ID записи",
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
        command=ex2
    )

    head.pack(fill='x')
    entry.pack(pady=5)
    searchB.pack(pady=5)
    backB.pack(pady=5)
    rootB.iconbitmap('assets/rok.ico')
    rootB.mainloop()

def create():
    global c1
    global P_id, dd, treat_1, price, med, ddd

    b1 = P_id.get()
    b2 = dd.get()
    b3 = treat_1.get()
    b4 = price.get()
    b5 = med.get()
    b6 = ddd.get()

    c1.execute("INSERT INTO TRAINING VALUES(?,?,?,?,?,?)", (b1, b2, b3, b4, b5, b6))
    conn.commit()

    c2 = conn.cursor()
    t = c2.execute('SELECT * FROM EPLOYEE where ENPLOYEE_ID=?', (b1,))

    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_date = datetime.now().strftime("%Y")
    first_date = datetime.now().strftime("%d-%m-%Y")

    context = {'year': file_date,
               'place': b5,
               'programm': b3,
               'first_date': first_date,
               'second_date': b4}
    for i in t:
        context['fullname'] = i[1]
        context['phone'] = i[5]
    doc = DocxTemplate("KvalShablon.docx")
    doc.render(context)

    output_folder = "KvalFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_name = f"generated_{current_time}.docx"
    file_path = os.path.join(output_folder, file_name)
    doc.save(file_path)

    tkinter.messagebox.showinfo("ROKARM система", "Запись о повышении квалификации успешно добавлена")


def exit():
    rootB.destroy()


def ex1():
    rootB.destroy()
    B_display()


def ex2():
    rootB.destroy()
    BILLING()


def ex3():
    rootDE.destroy()
    BILLING()


def ex4():
    rootB.destroy()
    delo()


def delling():
    global d1, de
    de = str(d1.get())
    p = list(conn.execute("select * from TRAINING where DOCUMENT_ID=?", (de,)))
    if len(p) != 0:
        conn.execute("DELETE from TRAINING where DOCUMENT_ID=?", (de,))
        dme = tkinter.Label(
            rootDE,
            text="Запись успешна удалена из базы",
            bg="#201E1F",
            fg="green"
        )

        dme.pack(fill='x')
        conn.commit()
    else:
        error = tkinter.Label(
            rootDE,
            text="Запись не найдена",
            bg="#201E1F",
            fg="Red"
        )

        error.pack(fill='x')


rootDE = None


def get_next_training_id():
    c = conn.cursor()
    c.execute("SELECT MAX(DOCUMENT_ID) FROM TRAINING")
    max_id = c.fetchone()[0]
    return max_id + 1 if max_id else 1


def BILLING():
    global rootB, L1, L3, treat1, P_id, dd, cost, med, med_q, price, treat_1, j, jjj, dd_d, ddd

    rootB = customtkinter.CTk()
    rootB.geometry("520x700+60+190")
    rootB.title("Повышение квалификации")

    next_id = get_next_training_id()

    head = tkinter.Label(
        rootB,
        text="Повышение квалификации",
        bg="#1A1919",
        fg='#daffda',
        font="Times 16 bold",
        padx=20,
        pady=10
    )

    head.pack(fill='x')

    id = customtkinter.CTkLabel(
        rootB,
        text="ID Сотрудника"
    )
    id.place(x=10, y=50)

    P_id = customtkinter.CTkEntry(
        rootB,
        width=230
    )
    P_id.place(x=10, y=80)

    dd_l = customtkinter.CTkLabel(
        rootB,
        text="№ Записи"
    )
    dd_l.place(x=10, y=110)

    dd = customtkinter.CTkEntry(
        rootB,
        width=230
    )
    dd.insert(0, next_id)
    dd.place(x=10, y=140)

    treat = customtkinter.CTkLabel(
        rootB,
        text="Дополнительные сведения"
    )
    treat.place(x=10, y=170)

    treat_1 = CTkListbox(
        rootB,
        width=230
    )

    dd_d = customtkinter.CTkLabel(
        rootB,
        text="Вид повышения"
    )
    dd_d.place(x=280, y=110)

    ddd = customtkinter.CTkEntry(
        rootB,
        width=230
    )
    ddd.place(x=280, y=140)

    treat_1.insert(0, "Очное краткосрочное")
    treat_1.insert(1, "Очное среднесрочное")
    treat_1.insert(2, "Очное долгосрочное")
    treat_1.insert(3, "Заочное краткосрочное")
    treat_1.insert(4, "Заочное среднесрочное")
    treat_1.insert(5, "Заочное долгосрочное")
    treat_1.insert(6, "Дистанционное краткосрочное")
    treat_1.insert(7, "Дистанционное среднесрочное")
    treat_1.insert("END", "Дистанционное долгосрочное")
    treat_1.place(x=10, y=200)

    costl = customtkinter.CTkLabel(
        rootB,
        text="Дата окончания (ГГГГ-ММ-ДД)"
    )
    costl.place(x=280, y=50)

    price = customtkinter.CTkEntry(
        rootB,
        width=230
    )
    price.place(x=280, y=80)

    med1 = customtkinter.CTkLabel(
        rootB,
        text="Наименование учреждения"
    )
    med1.place(x=280, y=170)

    med = CTkListbox(
        rootB,
        height=50
    )

    med.insert(0, "Корпорация Майкрософт")
    med.insert(1, "Accenture")
    med.insert(2, "Cognizant")
    med.insert(3, "Infosys")
    med.insert(4, "Tata Consultancy Services")
    med.insert(5, "SAP")
    med.insert(6, "Capgemini")
    med.insert(7, "IBM")
    med.insert(8, "Oracle")
    med.insert(9, "Deloitte")
    med.insert("END", "DXC")
    med.place(x=280, y=200)

    b1 = customtkinter.CTkButton(
        rootB,
        text="Создать",
        command=create,
        width=495
    )
    b1.place(x=10, y=430)

    SEARCH_N = customtkinter.CTkButton(
        rootB,
        text="Поиск",
        width=495,
        command=ex1
    )

    SEARCH_N.place(x=10, y=470)

    DEL_N = customtkinter.CTkButton(
        rootB,
        text="Удалить",
        width=495,
        command=ex4
    )

    DEL_N.place(x=10, y=510)

    graph = customtkinter.CTkButton(
        rootB,
        text="График сотрудников по организациям",
        width=495,
        command=plot_training_positions
    )

    graph.place(x=10, y=550)

    export_button = customtkinter.CTkButton(
        rootB,
        text="Экспорт в Excel",
        command=export_to_excel,
        width=495
    )
    export_button.place(x=10, y=590)

    ee = customtkinter.CTkButton(
        rootB,
        text="Выйти",
        command=exit,
        width=495
    )
    ee.place(x=10, y=630)

    rootB.iconbitmap('assets/rok.ico')
    rootB.mainloop()


def delo():
    global rootDE, d1
    rootDE = customtkinter.CTk()
    rootDE.geometry("360x610+210+130")
    rootDE.title("Удаление записи")
    rootDE.configure(bg="#201E1F")
    h1 = tkinter.Label(
        rootDE,
        text="Введите № документа\nдля удаления из базы",
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
        command=ex3,
        width=180,
    )

    h1.pack(fill='x')
    d1.pack(pady=5)
    B1.pack(pady=5)
    B2.pack(pady=5)
    rootDE.iconbitmap('assets/rok.ico')
    rootDE.mainloop()
