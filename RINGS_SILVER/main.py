import pandas as pd
import shutil
import os
from tkinter import messagebox
from tkinter import *


def vk_url_parser(vk_photos_html, main_number):
    photos = vk_photos_html.split("body>")[1][:-2]
    photos_list = photos.split("<hr />")[:-1]

    photo_list = list()
    for i in photos_list:
        photo_list.append(i.split('<img src="')[1].split('" /><div c')[0])
    main_photo = photo_list.pop(main_number - 1)
    return '\n'.join(photo_list), main_photo


def make_file():
    photos, main_photo = vk_url_parser(HTML_input.get(), int(photo_number.get()))
    name = name_input.get().strip()
    price = int(cost_input.get()) * 1.25
    weight = int(weight_input.get())
    annotation = annotation_input.get().strip()

    EXCEL_FILE_NAME: str = 'Серебро кольцо печатка мужское.xlsx'
    ARTICLE_FILE_NAME: str = 'article.txt'
    weight_in_box = 30 + weight

    df = pd.read_excel("EXAMPLE " + EXCEL_FILE_NAME, sheet_name='Шаблон для поставщика', header=[0,1,2])

    if os.path.exists(ARTICLE_FILE_NAME):
        with open(ARTICLE_FILE_NAME) as f:
            article = f.read()
            f = filter(str.isdecimal, article.split('ОК')[1])
            clear_digit = "".join(f)
            article = 'ОК' + str(int(clear_digit) + 1)

    df[df.columns[28]] = df.index + 1
    df[df.columns[1]] = article + df[df.columns[28]].astype(str)
    df[df.columns[28]] = None
    df[df.columns[3]] = price
    df[df.columns[9]] = weight_in_box
    df[df.columns[13]] = main_photo
    df[df.columns[14]] = photos
    df[df.columns[2]] = name + " " + df[df.columns[21]].astype(str)
    df[df.columns[19]] = name
    df[df.columns[25]] = name
    df[df.columns[27]] = annotation
    df[df.columns[26]] = weight

    df.reset_index()
    df.drop(columns=df.columns[0], axis=1, inplace=True)

    shutil.copyfile("EXAMPLE " + EXCEL_FILE_NAME, article + ".xlsx")

    df.to_csv("temp.csv", index=False)
    df = pd.read_csv("temp.csv")
    with pd.ExcelWriter(article + ".xlsx", engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name='Шаблон для поставщика', index=False)

    with open(ARTICLE_FILE_NAME, "w") as file:
        file.write(article)

    messagebox.showinfo(title="Сообщение", message=f"Файл {article}.xlsx \n успешно создан")
    name_input.delete(0, END)
    weight_input.delete(0, END)
    cost_input.delete(0, END)
    annotation_input.delete(0, END)
    HTML_input.delete(0, END)
    photo_number.delete(0, END)


if __name__ == "__main__":
    root = Tk()

    window_height = 910
    window_width = 200
    root.title("Генерация файлов для OZON")
    root.geometry(f"{window_height}x{window_width}")

    root.resizable(width=False, height=False)

    canvas = Canvas(root, height=window_height, width=window_width)
    canvas.pack()

    frame = Frame(root, bg="gray")
    frame.place(relwidth=1, relheight=1)

    title1 = Label(frame, text="Название", bg="white")
    title1.grid(row=0, column=0, padx=10, pady=10)
    name_input = Entry(frame, bg="white", width=50)
    name_input.grid(row=1, column=0)

    title2 = Label(frame, text="Вес", bg="white")
    title2.grid(row=2, column=0, padx=10, pady=10)
    weight_input = Entry(frame, bg="white")
    weight_input.grid(row=3, column=0)

    title3 = Label(frame, text="Цена", bg="white")
    title3.grid(row=0, column=1, padx=10, pady=10)
    cost_input = Entry(frame, bg="white")
    cost_input.grid(row=1, column=1)

    title4 = Label(frame, text="Аннотация", bg="white")
    title4.grid(row=2, column=1, padx=10, pady=10)
    annotation_input = Entry(frame, bg="white", width=50)
    annotation_input.grid(row=3, column=1)

    title5 = Label(frame, text="Исходный код страницы с фото", bg="white")
    title5.grid(row=0, column=2, padx=10, pady=10)
    HTML_input = Entry(frame, bg="white", width=50)
    HTML_input.grid(row=1, column=2)

    title6 = Label(frame, text="Номер основного фото", bg="white")
    title6.grid(row=2, column=2, padx=10, pady=10)
    photo_number = Entry(frame, bg="white")
    photo_number.grid(row=3, column=2)

    btn = Button(frame, text="Сгенерировать", bg="red", command=make_file)
    btn.grid(row=15, column=1, padx=10, pady=30)

    root.mainloop()