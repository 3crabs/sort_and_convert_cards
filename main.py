import csv
import functools
import tkinter.ttk
from threading import Thread
from tkinter import *
from tkinter import filedialog as fd

template_file_path = "../resources/Шаблон.xlsx"
window = Tk()
window.eval('tk::PlaceWindow . center')
file_path = StringVar()
dir_path = StringVar()
start_button = {}
message = StringVar()
message_label = None
selected_set = StringVar()
string_of_sets = StringVar()

rows = []

color_map = {
    "W": [0, "\nWhite\n"],
    "U": [1, "\nBlue\n"],
    "B": [2, "\nBlack\n"],
    "R": [3, "\nRed\n"],
    "G": [4, "\nGreen\n"],
    "Z": [5, "\nColorless / Lands\n"],
    "mc": [6, "\nMulticolor\n"]
}


class Row:

    def __init__(self, amount, eng_name, rus_name, card_set, condition, lang, foil, price, color):
        self.amount = amount
        self.eng_name = eng_name
        self.rus_name = rus_name
        self.card_set = card_set
        self.condition = condition
        self.lang = lang
        self.foil = foil
        self.price = price
        self.color = color

    def __str__(self) -> str:
        return f'{self.amount} {self.eng_name} {self.rus_name} {self.card_set} ' \
               f'{self.condition} {self.lang} {self.foil} {self.price}'

    def format_output(self):
        rus_name = " "
        foil = " "

        if self.foil == "1":
            foil = " FOIL "

        if self.rus_name != "---":
            rus_name = f" / {self.rus_name} "
        return f'{self.amount} {self.eng_name}{rus_name}{self.card_set} ' \
               f'{self.condition} {self.lang}{foil}{self.price} \n'


def get_color_value(color):
    if color == "":
        color = 'Z'
    if len(color) > 1:
        color = "mc"
    return color_map[color]


def check_button_state():
    global start_button
    start_button["state"] = NORMAL
    if file_path.get() == '':
        start_button["state"] = DISABLED
    if dir_path.get() == '':
        start_button["state"] = DISABLED


def select_dir():
    global dir_path
    dir_name = fd.askdirectory(title='Выберите папку для сохранения результатов')
    dir_path.set(dir_name)
    check_button_state()


def select_file():
    global file_path

    # file_name = fd.askopenfile(title='Выберите файл данных', filetypes=[('xlsx files', ['.xlsx'])])
    file_name = fd.askopenfile(title='Выберите файл данных', filetypes=[('csv files', ['.csv'])])
    file_path.set(file_name.name)
    check_button_state()


def sort_cards(first_row, second_row):
    # цвет -> цена -> eng_name
    if first_row.color[0] > second_row.color[0]:
        return 1
    elif first_row.color[0] == second_row.color[0]:
        if first_row.price < second_row.price:
            return 1
        elif first_row.price == second_row.price:
            if first_row.eng_name > second_row.eng_name:
                return 1
    return -1


def work():
    global message

    progress_bar = tkinter.ttk.Progressbar(
        window,
        orient='horizontal',
        mode='indeterminate',
        length=450,
    )
    progress_bar.grid(row=5, column=0, columnspan=3, padx=5, pady=5)
    progress_bar.start(10)

    # Залупа по экселю
    # workbook = load_workbook(file_path.get())
    # ws = workbook.worksheets[0]
    # i = 2
    # while ws[i][0].value is not None:
    #     color = ws[i][4].value
    #     if color is None:
    #         color = 'Z'
    #     r = Row(ws[i][10].value, ws[i][0].value, ws[i][1].value,
    #             ws[i][3].value, ws[i][7].value, ws[i][6].value,
    #             ws[i][8].value, ws[i][11].value, color)
    #     rows.append(r)
    #     i += 1
    #     print(i)

    if len(rows) == 0:
        with open(file_path.get(), encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=";")
            next(reader)
            useless_sets = string_of_sets.get().split(',')
            for row in reader:
                if row[3] not in useless_sets:
                    r = Row(row[10], row[0], row[1],
                            row[3], row[7], row[6],
                            row[8], int(row[11]), get_color_value(row[4]))
                    rows.append(r)

        rows.sort(key=functools.cmp_to_key(sort_cards))

    if selected_set.get():
        one_set()
    else:
        all_sets()

    progress_bar.stop()
    progress_bar.grid_remove()
    check_button_state()


def one_set():
    with open(dir_path.get() + r'/one_set_cards.txt', 'w', encoding='utf-8') as txt_file:
        prev_card_color = None
        txt_file.write(f"{selected_set.get()}\n")
        for i in range(len(rows)):
            if rows[i].card_set == selected_set.get():
                if rows[i].color[1] != prev_card_color:
                    prev_card_color = rows[i].color[1]
                    txt_file.write(f"{rows[i].color[1]}\n")
                txt_file.write(rows[i].format_output())
        message.set('Готово!')


def all_sets():
    with open(dir_path.get() + r'/cards.txt', 'w', encoding='utf-8') as txt_file:
        prev_card_set = None
        prev_card_color = None
        for i in range(len(rows)):
            if rows[i].color[1] != prev_card_color:
                prev_card_color = rows[i].color[1]
                txt_file.write(f"{rows[i].color[1]}\n")

            txt_file.write(rows[i].format_output())
        message.set('Готово!')


def start():
    global file_path
    global dir_path
    global message

    start_button["state"] = DISABLED
    message.set('Обрабротка начата')
    thread = Thread(target=work)
    thread.start()


def main():
    global window
    global file_path
    global dir_path
    global start_button
    global message
    global message_label
    global selected_set
    global string_of_sets

    window.title("Нормально делаем")
    window.geometry('350x180')

    file_label = Label(text="Путь к файлу:")
    file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
    file_input = Entry(textvariable=file_path)
    file_input.grid(row=0, column=1, padx=5, pady=5)
    file_button = Button(window, text="Выбрать файл", command=select_file)
    file_button.grid(column=2, row=0, padx=5, pady=5, sticky="e")

    dir_label = Label(text="Путь к папке:")
    dir_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
    dir_input = Entry(textvariable=dir_path)
    dir_input.grid(row=1, column=1, padx=5, pady=5)
    dir_button = Button(window, text="Выбрать папку", command=select_dir)
    dir_button.grid(column=2, row=1, padx=5, pady=5, sticky="e")

    message_label = Label(textvariable=message)
    message_label.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="w")
    message.set("Заполните поля, нажмите кнопку."[0:35])

    start_button = Button(window, text="Начать", command=start)
    start_button.grid(column=2, row=4, padx=5, pady=5, sticky="e")
    start_button["state"] = DISABLED

    selected_set_label = Label(text="Введите сет:")
    selected_set_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

    select_set_input = Entry(textvariable=selected_set)
    select_set_input.grid(row=2, column=1, padx=5, pady=5)

    string_of_sets_label = Label(text="Ненужные сеты:")
    string_of_sets_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")

    string_of_sets_input = Entry(textvariable=string_of_sets)
    string_of_sets_input.grid(row=3, column=1, padx=5, pady=5)

    window.mainloop()


if __name__ == '__main__':
    main()
