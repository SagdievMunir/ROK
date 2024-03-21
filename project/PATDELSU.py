import tkinter
import sqlite3
import tkinter.messagebox
import customtkinter
import pandas as pd
import os
from datetime import datetime

conn = sqlite3.connect("database/ROK.db")

rootU = None
rootD = None
rootS = None
head = None
inp_s = None
searchB = None

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


def Search_button():
    global inp_s, entry, errorS, t, i, q, dis1, dis2, dis3, dis4, dis5, dis6, dis7, dis8, dis9, dis10
    global l1, l2, l3, l4, l5, l6, l7, l8, l9, l10
    c1 = conn.cursor()
    inp_s = entry.get()

    p = list(c1.execute('select * from ENCOURAGEMENT where DOCUMENT_ID=?', (inp_s,)))

    if len(p) == 0:
        if 'errorS' in globals():
            errorS.destroy()
        errorS = tkinter.Label(
            rootS,
            text="Данные о поощрении не найдены",
            bg="#1A1919",
            fg='#daffda',
            font="Times 16 bold"
        )
        errorS.pack(fill='x')
    else:
        if 'errorS' in globals():
            errorS.destroy()
        t = c1.execute('SELECT * FROM ENCOURAGEMENT where DOCUMENT_ID=?', (inp_s,))

        # Удаление предыдущих меток
        for widget in rootS.winfo_children():
            if isinstance(widget, tkinter.Label):
                widget.destroy()

        for i in t:
            l2 = tkinter.Label(
                rootS,
                text="ID сотрудника",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis2 = tkinter.Label(
                rootS,
                text=i[0],
                fg='white',
                bg="#1A1919",
            )

            l3 = tkinter.Label(
                rootS,
                text="ID документа",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis3 = tkinter.Label(
                rootS,
                fg='white',
                bg="#1A1919",
                text=i[1]
            )

            l4 = tkinter.Label(
                rootS,
                text="Мотив поощрения",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis4 = tkinter.Label(
                rootS,
                fg='white',
                bg="#1A1919",
                text=i[2],
            )

            l5 = tkinter.Label(
                rootS,
                text="Тип поощрения",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis5 = tkinter.Label(
                rootS,
                fg='white',
                bg="#1A1919",
                text=i[3]
            )

            l6 = tkinter.Label(
                rootS,
                text="Дата поощрения",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis6 = tkinter.Label(
                rootS,
                fg='white',
                bg="#1A1919",
                text=i[4]
            )

            l7 = tkinter.Label(
                rootS,
                text="Основание",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis7 = tkinter.Label(
                rootS,
                fg='white',
                bg="#1A1919",
                text=i[5]
            )

            l8 = tkinter.Label(
                rootS,
                text="Сумма",
                fg='#daffda',
                bg="#1A1919",
                font="Times 12 bold"
            )

            dis8 = tkinter.Label(
                rootS,
                fg='white',
                bg="#1A1919",
                text=i[6]
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
            conn.commit()



def eXO():
    rootS.destroy()


def eX1():
    rootD.destroy()


def P_display():
    global rootS, head, inp_s, entry, searchB
    rootS = customtkinter.CTk()
    rootS.geometry('360x590+1350+130')
    rootS.title("Окно поиска")

    head = tkinter.Label(
        rootS,
        text="Введите № документа",
        bg="#1A1919",
        fg='#daffda',
        font=("Times", 16, "bold"),
        padx=20,
        pady=10
    )

    entry = customtkinter.CTkEntry(
        rootS,
        width=180,
    )

    searchB = customtkinter.CTkButton(
        rootS,
        text='Поиск',
        width=180,
        command=Search_button
    )

    backB = customtkinter.CTkButton(
        rootS,
        text='Выйти',
        width=180,
        command=eXO
    )

    menubar = tkinter.Menu(rootS)
    filemenu = tkinter.Menu(menubar, tearoff=0)
    filemenu.add_command(label="Новый", command=P_display)
    filemenu.add_separator()
    filemenu.add_command(label="Выйти", command=eXO)
    menubar.add_cascade(label="Файл", menu=filemenu)
    rootS.config(menu=menubar)
    head.pack(fill='x')
    entry.pack(pady=5)
    searchB.pack(pady=5)
    backB.pack(pady=5)
    rootS.iconbitmap('assets/rok.ico')
    rootS.mainloop()


inp_d = None
entry1 = None
errorD = None
disd1 = None


def Delete_button():
    global inp_d, entry1, errorD, disd1
    # c1 = conn.cursor()
    inp_d = entry1.get()
    p = list(conn.execute("select * from ENCOURAGEMENT where DOCUMENT_ID=?", (inp_d,)))
    if len(p) == 0:
        errorD = tkinter.Label(
            rootD,
            text="Информация о поощрении не найдена",
            bg="#1A1919",
            fg='#daffda',
            font="Times 16 bold"
        )
        errorD.pack(fill='x')
    else:
        conn.execute('DELETE FROM ENCOURAGEMENT where DOCUMENT_ID=?', (inp_d,))
        disd1 = tkinter.Label(
            rootD,
            text="Информация о поощрении удалена",
            bg="#1A1919",
            fg='#daffda',
            font="Times 16 bold",
            padx=20,
            pady=10
        )
        disd1.pack(fill='x')
        conn.commit()


def D_display():
    global rootD, headD, inp_d, entry1, DeleteB
    rootD = customtkinter.CTk()
    rootD.geometry('360x590+1350+130')
    rootD.title("Окно удаления")

    headD = tkinter.Label(
        rootD,
        text="Введите ID сотрудника",
        bg="#1A1919",
        fg='#daffda',
        font="Times 16 bold",
        padx=20,
        pady=10
    )

    entry1 = customtkinter.CTkEntry(
        rootD,
        width=180,
    )

    DeleteB = customtkinter.CTkButton(
        rootD,
        width=180,
        text="Удалить",
        command=Delete_button
    )

    BackB = customtkinter.CTkButton(
        rootD,
        width=180,
        text="Выйти",
        command=eX1
    )

    headD.pack(fill='x')
    entry1.pack(pady=5)
    DeleteB.pack(pady=5)
    BackB.pack(pady=5)
    rootD.iconbitmap('assets/rok.ico')
    rootD.mainloop()


pat_ID, pat_name, pat_dob, pat_address, pat_sex, pat_BG, pat_email = None, None, None, None, None, None, None
pat_contact, pat_contactalt, pat_CT = None, None, None
u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, ue1 = None, None, None, None, None, None, None, None, None, None, None


def up1():
    global u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, ue1, conn
    conn.cursor()
    u1 = pat_ID.get()
    u2 = pat_name.get()
    u3 = pat_sex.get()
    u4 = pat_dob.get()
    u5 = pat_BG.get()
    u6 = pat_contact.get()
    u7 = pat_contactalt.get()
    u8 = pat_email.get()
    conn = sqlite3.connect("database/ROK.db")
    p = list(conn.execute("Select * from ENCOURAGEMENT where DOCUMENT_ID=?", (u1,)))
    if len(p) != 0:
        conn.execute(
            'UPDATE ENCOURAGEMENT SET ENPLOYEE_ID=?,'
            'MOTIVE=?,TYPE_ENCOURAGEMENT=?,DATE=?,EXPLANATION=?,AMOUNT=? where DOCUMENT_ID=?',
            (u2, u3, u4, u5, u6, u7, u8))
        tkinter.messagebox.showinfo("ROK ARM система", "Информация о сотруднике успешно обновлена")
        conn.commit()

    else:
        tkinter.messagebox.showinfo("ROK ARM система", "Сотрудник не зарегистрирован")


labelu = None
bu1 = None


def EXITT():
    rootU.destroy()


regform, id_new, name, dob, sex, email, ct, addr = None, None, None, None, None, None, None, None
c1, c2, bg, SUBMIT, menubar, filemenu, p1f, p2f, HEAD = None, None, None, None, None, None, None, None, None


def P_UPDATE():
    global pat_address, pat_BG, pat_contact, pat_contactalt, pat_CT, pat_dob, pat_email, pat_ID, pat_name, pat_sex
    global rootU, regform, id_new, name, dob, sex, email, ct, addr, c1, c2, bg, SUBMIT, menubar, filemenu, p1f, p2f, HEAD
    rootU = customtkinter.CTk()
    rootU.title("Обновление данных")
    rootU.geometry('380x600+1350+130')

    menubar = tkinter.Menu(rootU)
    filemenu = tkinter.Menu(menubar, tearoff=0)
    filemenu.add_command(label="Новый", command=P_UPDATE)
    filemenu.add_separator()
    filemenu.add_command(label="Выйти", command=EXITT)
    rootU.config(menu=menubar)
    menubar.add_cascade(label="Файл", menu=filemenu)

    HEAD = tkinter.Label(
        rootU,
        text="Введите информацию для обновления",
        fg="#daffda",
        bg="#1A1919",
        font="Times 16 bold",
        padx=20,
        pady=10
    )

    id_new = customtkinter.CTkLabel(
        rootU,
        width=180,
        text="ID Записи"
    )

    pat_ID = customtkinter.CTkEntry(
        rootU,
        width=180,
    )

    name = customtkinter.CTkLabel(
        rootU,
        width=180,
        text="ID Сотрудника"
    )

    pat_name = customtkinter.CTkEntry(
        rootU,
        width=180,
    )

    sex = customtkinter.CTkLabel(
        rootU,
        width=180,
        text="Мотив поощрения"
    )

    pat_sex = customtkinter.CTkEntry(
        rootU,
        width=180,
    )

    dob = customtkinter.CTkLabel(
        rootU,
        width=180,
        text="Дата (ГГГГ-ММ-ДД)"
    )

    pat_dob = customtkinter.CTkEntry(
        rootU,
        width=180,
    )

    bg = customtkinter.CTkLabel(
        rootU,
        width=180,
        text="Вид поощрения"
    )

    pat_BG = customtkinter.CTkEntry(
        rootU,
        width=180,
    )

    c1 = customtkinter.CTkLabel(
        rootU,
        width=180,
        text="Номер документа"
    )

    pat_contact = customtkinter.CTkEntry(
        rootU,
        width=180,
    )

    c2 = customtkinter.CTkLabel(
        rootU,
        width=180,
        text="Основание"
    )

    pat_contactalt = customtkinter.CTkEntry(
        rootU,
        width=180,
    )

    email = customtkinter.CTkLabel(
        rootU,
        width=180,
        text="Сумма"
    )

    pat_email = customtkinter.CTkEntry(
        rootU,
        width=180,
    )

    SUBMIT = customtkinter.CTkButton(
        rootU,
        width=180,
        text="Подтвердить",
        command=up1
    )

    ex = customtkinter.CTkButton(
        rootU,
        width=180,
        text="Выйти",
        command=EXITT
    )

    HEAD.pack(fill='x')
    id_new.pack()
    pat_ID.pack()
    name.pack()
    pat_name.pack()
    sex.pack()
    pat_sex.pack()
    dob.pack()
    pat_dob.pack()
    bg.pack()
    pat_BG.pack()
    c1.pack()
    pat_contact.pack()
    c2.pack()
    pat_contactalt.pack()
    email.pack()
    pat_email.pack()
    SUBMIT.pack(pady=5)
    ex.pack()
    rootU.iconbitmap('assets/rok.ico')
    rootU.mainloop()
