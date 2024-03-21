import tkinter

import customtkinter

from mainmenu import menu


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

# variables
root = None
userbox = None
passbox = None
topframe = None
bottomframe = None
frame3 = None
login = None
# error = None


def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    window.geometry('{}x{}+{}+{}'.format(width, height, x, y))


def GET():
    global userbox, passbox, error, root

    S1 = userbox.get()
    S2 = passbox.get()

    if S1 == 'Admin' and S2 == '1234':
        root.withdraw()
        root.quit()
        menu(center_window)
    elif S1 == '' and S2 == '':
        root.withdraw()
        root.quit()
        menu(center_window)
    else:
        if 'error' in globals():
            error.destroy()
        error = tkinter.Label(
            bottomframe,
            text="Неправильный логин или пароль \n Попробуйте снова",
            fg="red",
            bg="#201E1F",
            font="bold"
        )
        error.place(x=105, y=305)


def Entry():
    global userbox, passbox, login, topframe, bottomframe, root

    root = customtkinter.CTk()
    root.geometry("500x300")
    root.configure(fg_color="#201E1F")
    root.resizable(width=False, height=False)

    heading = tkinter.Label(
        root,
        text="Рабочая информационная система \n ARM ROK",
        fg="#daffda",
        font=("Times", 16, "bold"),
        bg="#1A1919",
        padx=20,
        pady=10
    )

    username = customtkinter.CTkLabel(
        root,
        width=180,
        font=("Times", 18,),
        text="Логин"
    )

    userbox = customtkinter.CTkEntry(
        root,
        width=180,
        show="*"
    )

    password = customtkinter.CTkLabel(
        root,
        width=180,
        font=("Times", 18,),
        text="Пароль"
    )

    passbox = customtkinter.CTkEntry(
        root,
        width=180,
        show="*"
    )

    login = customtkinter.CTkButton(
        root,
        width=180,
        font=("Times", 16,),
        text="Авторизироваться",
        command=GET
    )

    heading.pack(fill='x')
    username.place(x=135, y=80)
    userbox.place(x=135, y=110)
    password.place(x=135, y=160)
    passbox.place(x=135, y=190)
    login.place(x=135, y=270)

    root.title("Авторизация в системе")

    center_window(root, 450, 350)
    root.iconbitmap('assets/rok.ico')
    root.mainloop()


Entry()
