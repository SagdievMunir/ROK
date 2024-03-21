import sqlite3
import tkinter
import tkinter.messagebox
import customtkinter
from CTkListbox import *
from docxtpl import DocxTemplate
from datetime import datetime, timedelta
import os
import matplotlib.pyplot as plt
import pandas as pd

conn = sqlite3.connect("database/ROK.db")

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

P_id = None
rootR = None

def export_to_excel():
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_name = f"vacation_{current_time}.xlsx"

    output_folder = "ExcelFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_path = os.path.join(output_folder, file_name)

    df = pd.read_sql_query("SELECT * FROM VACATION", conn)
    df.to_excel(file_path, index=False)

    tkinter.messagebox.showinfo("ROKARM система", "Данные экспортированы в таблицу")


def plot_vacation_positions():
    c = conn.cursor()

    c.execute("SELECT VACATION_TYPE FROM VACATION")
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
    plt.xlabel('Вид отпуска', color='white')
    plt.ylabel('Количество сотрудников', color='white')
    plt.title('Количество отпусков по видам', color='white')
    plt.xticks(rotation=45, ha='right')
    plt.grid(linewidth=0.2)
    plt.tick_params(axis='x', colors='white')
    plt.tick_params(axis='y', colors='white')
    plt.yticks(list(range(int(min(counts)), int(max(counts)) + 1, 1)))
    plt.tight_layout()
    plt.get_current_fig_manager().window.geometry('+600+300')
    plt.show()


def delling():
    global d1, de
    de = str(d1.get())
    p = list(conn.execute("select * from VACATION where DOCUMENT_ID=?", (de,)))
    if len(p) != 0:
        conn.execute("DELETE from VACATION where DOCUMENT_ID=?", (de,))
        dme = tkinter.Label(
            rootDE,
            text="Отпуск успешно удален из базы",
            bg="#201E1F",
            fg="green"
        )

        dme.pack(fill='x')
        conn.commit()
    else:
        error = tkinter.Label(
            rootDE,
            text="Отпуск не найден",
            bg="#201E1F",
            fg="Red"
        )

        error.pack(fill='x')


rootDE = None


def get_next_vacation_id():
    c = conn.cursor()
    c.execute("SELECT MAX(DOCUMENT_ID) FROM VACATION")
    max_id = c.fetchone()[0]
    return max_id + 1 if max_id else 1


##ROOM BUTTON
def room_button():
    global P_id, r1, r2, room_t, da, dd, rate, room_no, r3, r4, r5, r6, conn

    r1 = P_id.get()
    r2 = room_t.get()
    r3 = room_no.get()
    r4 = rate.get()
    r5 = da.get()
    r6 = dd.get()
    conn.execute('INSERT INTO VACATION VALUES(?,?,?,?,?,?)', (r1, r3, r2, r4, r5, r6,))

    c1 = conn.cursor()
    t = c1.execute('SELECT * FROM EPLOYEE where ENPLOYEE_ID=?', (r3,))

    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_date = datetime.now().strftime("%d.%m.%Y")

    context = {'first_date': file_date,
               'third_date': r4,
               'fourth_date': (datetime.strptime(r4, "%Y-%m-%d") + timedelta(days=int(r5))).date(),
               'type_vacation': r2}
    for i in t:
        context['fullname'] = i[1]
        context['post'] = i[4]
        context['wages'] = i[8]
    doc = DocxTemplate("OtpuskShablon.docx")
    doc.render(context)

    output_folder = "OtpuskFiles"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_name = f"generated_{current_time}.docx"
    file_path = os.path.join(output_folder, file_name)
    doc.save(file_path)

    tkinter.messagebox.showinfo("ROKARM система", "Отпуск назначен")
    conn.commit()


def update_button():
    global P_id, r1, r2, room_t, da, dd, rate, room_no, r3, r4, r5, r6, conn
    r1 = P_id.get()
    r2 = room_t.get()
    r3 = room_no.get()
    r4 = rate.get()
    r5 = da.get()
    r6 = dd.get()
    p = list(conn.execute("Select * from VACATION where DOCUMENT_ID=?", (r1,)))
    if len(p) != 0:
        conn.execute(
            'UPDATE VACATION SET DOCUMENT_ID=?,VACATION_TYPE=?,DATE=?,NUMBER_OF_DAYS=?,EXPLANATION=? where '
            'DOCUMENT_ID=?',
            (r3, r2, r4, r5, r6, r1,))
        tkinter.messagebox.showinfo("ROKARM система", "Данные об отпуске обновлены")
        conn.commit()
    else:
        tkinter.messagebox.showinfo("ROKARM система", "Нет данных по отпусках")


##ROOT FOR DISPLAY ROOM INFO
rootRD = None


##EXIT FOR ROOM_PAGE
def EXITT():
    global rootR
    rootR.destroy()


##FUNCTION FOR ROOM DISPLAY BUTTON
def ROOMD_button():
    global errorS, c1, conn, P_iid
    c1 = conn.cursor()
    r1 = P_iid.get()

    p = list(c1.execute('select * from VACATION where DOCUMENT_ID=?', (r1,)))

    if len(p) == 0:
        if 'errorS' in globals():
            errorS.destroy()
        errorS = tkinter.Label(
            rootRD,
            text="Данные о записи не найдены",
            bg="#1A1919",
            fg='#daffda',
            font="Times 16 bold"
        )
        errorS.pack(fill='x')
    else:
        t = c1.execute('SELECT * FROM VACATION where DOCUMENT_ID=?', (r1,))

        # Удаление предыдущих меток
        for widget in rootRD.winfo_children():
            if isinstance(widget, tkinter.Label):
                widget.destroy()

        for ii in t:
            lr0 = tkinter.Label(
                rootRD,
                text="ID Сотрудника",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )
            dis0 = tkinter.Label(
                rootRD,
                text=ii[0],
                fg='white',
                bg="#1A1919",
            )
            lr1 = tkinter.Label(
                rootRD,
                text="№ записи",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )
            dis1 = tkinter.Label(
                rootRD,
                text=ii[1],
                fg='white',
                bg="#1A1919",
            )
            lr2 = tkinter.Label(
                rootRD,
                text="Тип отпуска",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )
            dis2 = tkinter.Label(
                rootRD,
                fg='white',
                bg="#1A1919",
                text=ii[2]
            )
            lr3 = tkinter.Label(
                rootRD,
                text="Дата отпуска",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )
            dis3 = tkinter.Label(
                rootRD,
                fg='white',
                bg="#1A1919",
                text=ii[3]
            )
            lr4 = tkinter.Label(
                rootRD,
                text="Количество дней отпуска",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )
            dis4 = tkinter.Label(
                rootRD,
                fg='white',
                bg="#1A1919",
                text=ii[4]
            )

            lr0.pack(fill="x")
            dis0.pack(fill="x")
            lr1.pack(fill="x")
            dis1.pack(fill="x")
            lr2.pack(fill="x")
            dis2.pack(fill="x")
            lr3.pack(fill="x")
            dis3.pack(fill="x")
            lr4.pack(fill="x")
            dis4.pack(fill="x")

        conn.commit()



def exittt():
    rootRD.destroy()
    Room_all()

def exit1():
    rootR.destroy()
    roomDD()


def exit2():
    rootDE.destroy()
    Room_all()


def exit3():
    rootR.destroy()
    delo()


def roomDD():
    global rootRD, ra1, ss, P_iid
    rootRD = customtkinter.CTk()
    rootRD.geometry("440x440+135+340")
    rootRD.resizable(width=False, height=False)
    rootRD.title("Поиск по отпускам")

    ra1 = tkinter.Label(
        rootRD,
        text="Введите № записи",
        bg="#1A1919",
        fg='#daffda',
        font="Times 16 bold",
        padx=20,
        pady=10
    )

    ra1.pack(fill="x")

    P_iid = customtkinter.CTkEntry(
        rootRD,
        width=180,
    )

    P_iid.pack(pady=5)

    ss = customtkinter.CTkButton(
        rootRD,
        text="Поиск",
        width=180,
        command=ROOMD_button
    )

    ss.pack(pady=5)

    e = customtkinter.CTkButton(
        rootRD,
        text="Назад",
        width=180,
        command=exittt
    )

    e.pack(pady=5)

    rootRD.iconbitmap('assets/rok.ico')
    rootRD.mainloop()


def exitt():
    rootR.destroy()


L = None
L1 = None


def Room_all():
    global rootR, r_head, P_id, id, room_tl, L, i, room_t, room_nol, room_no, L1, j, ratel, rate, da_l, da, dd_l, dd, Submit, Update, cr
    rootR = customtkinter.CTk()
    rootR.title("Информация об отпусках")
    rootR.geometry("440x490+135+290")
    rootR.resizable(width=False, height=False)

    # Получение следующего доступного ID для отпуска
    next_id = get_next_vacation_id()

    r_head = tkinter.Label(
        rootR,
        text="Информация об отпусках",
        bg="#1A1919",
        fg='#daffda',
        font="Times 16 bold",
        padx=20,
        pady=10
    )

    r_head.place(x=0, y=0)
    r_head.pack(fill='x')

    id = customtkinter.CTkLabel(
        rootR,
        text="№ записи"
    )

    id.place(x=10, y=50)

    P_id = customtkinter.CTkEntry(
        rootR,
        width=230
    )
    P_id.insert(0, next_id)
    P_id.place(x=10, y=80)

    room_tl = customtkinter.CTkLabel(
        rootR,
        text="Вид отпуска"
    )

    room_tl.place(x=10, y=110)

    room_t = CTkListbox(
        rootR,
    )

    room_t.insert(0, "Основной")
    room_t.insert(1, "Дополнительный")
    room_t.insert(2, "Учебный без З/П")
    room_t.insert(3, "Учебный c З/П")
    room_t.insert(4, "Уход за ребёнком")
    room_t.insert(5, "По беременности")
    room_t.insert("END", "Неоплачиваемый")

    room_t.place(x=10, y=140)

    room_nol = customtkinter.CTkLabel(
        rootR,
        text="ID сотрудника"
    )

    room_nol.place(x=260, y=50)

    room_no = customtkinter.CTkEntry(
        rootR,
        width=160,
    )

    room_no.place(x=260, y=80)

    ratel = customtkinter.CTkLabel(
        rootR,
        text="Дата отпуска (ГГГГ-ММ-ДД)"
    )

    ratel.place(x=260, y=110)

    rate = customtkinter.CTkEntry(
        rootR,
        width=160,
    )

    rate.place(x=260, y=140)

    da_l = customtkinter.CTkLabel(
        rootR,
        text="Количество дней отпуска"
    )

    da_l.place(x=260, y=170)

    da = customtkinter.CTkEntry(
        rootR,
        width=160,
    )

    da.place(x=260, y=200)

    dd_l = customtkinter.CTkLabel(
        rootR,
        text="Основание"
    )

    dd_l.place(x=260, y=230)

    dd = customtkinter.CTkEntry(
        rootR,
        width=160,
    )

    dd.place(x=260, y=260)
    #dd.pack()

    Submit = customtkinter.CTkButton(
        rootR,
        text="Подтвердить",
        width=160,
        command=room_button
    )

    Submit.place(x=260, y=300)

    Update = customtkinter.CTkButton(
        rootR,
        text="Обновить",
        width=230,
        command=update_button
    )

    Update.place(x=10, y=410)

    cr = customtkinter.CTkButton(
        rootR,
        text='Поиск по отпускам',
        width=230,
        command=exit1
    )

    cr.place(x=10, y=370)

    e1 = customtkinter.CTkButton(
        rootR,
        text="Удалить",
        width=160,
        command=exit3
    )

    e1.place(x=260, y=340)

    ee = customtkinter.CTkButton(
        rootR,
        text="Выход",
        width=230,
        command=exitt
    )

    ee.place(x=10, y=450)

    plot_button = customtkinter.CTkButton(
        rootR,
        width=160,
        text="График сотрудников\n по отпускам",
        command=plot_vacation_positions
    )
    plot_button.place(x=260, y=380)

    export_button = customtkinter.CTkButton(
        rootR,
        width=160,
        text="Экспорт в Excel",
        command=export_to_excel
    )
    export_button.place(x=260, y=430)

    rootR.iconbitmap('assets/rok.ico')
    rootR.mainloop()


def delo():
    global rootDE, d1
    rootDE = customtkinter.CTk()
    rootDE.geometry("360x610+210+130")
    rootDE.title("Удаление отпуска")
    rootDE.configure(bg="#201E1F")
    h1 = tkinter.Label(
        rootDE,
        text="Введите № записи отпуска\nдля удаления из базы",
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
        command=exit2,
        width=180,
    )

    h1.pack(fill='x')
    d1.pack(pady=5)
    B1.pack(pady=5)
    B2.pack(pady=5)
    rootDE.iconbitmap('assets/rok.ico')
    rootDE.mainloop()