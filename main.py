import os
from datetime import datetime
from PyQt5.QtCore import QRegExp
from PyQt5.QtWidgets import (QApplication, QComboBox)
from PyQt5.uic import loadUiType
import openpyxl
import sys

Ui_MainWindow, QMainWindow = loadUiType("MainWindowUI.ui")


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.PBClose.clicked.connect(self.close)
        self.PBClear.clicked.connect(self.clearLineEditContents)

        #Заполняем дату
        self.DEDate.setDate(datetime.now().date())

        #Заполняем ответственых
        self.CBAgreed.addItems(["Начальник  участка- Крицкий С.Г.", "ИО Начальника  участка- Чащин Ю.М.",
                                "ИО Начальника  участка- Скурихин А. А."])

        #Заполняем направление куда ТМЦ
        self.CBRecipient.addItems(["Уч. Кувай", "г. Красноярск, ул. 26 Бакинских Комиссаров - Склад Б/У",
                                   "г. Красноярск, ул. 26 Бакинских Комиссаров", "Уч. Караган"])

        #Заполняем транспорт
        self.CBTransport.addItems(["MAN Г/н С615ЕР124",
                                   "Mitsubishi L200 Г/н С673ОВ124",
                                   "FUSO Г/н К700ХВ124",
                                   "MAN Г/н Х253ВО124 Прицеп Г/н МВ5912 24",
                                   "ISUZU RUSTRAC Г/н Т213СК124",
                                   "MAN Г/н С469ЕР124",
                                   "Камаз 65117-62  Г/н В775НВ124",
                                   "УАЗ Профи Г/н Н868РУ124",
                                   "НИВА Г/н Е776ТЕ124",
                                   "SCANIA Г/н С774СН 124 Прицеп Г/н МС9034 24",
                                   "MAN  Г/н М231ОЕ124",
                                   "Mitsubishi L200  Г/н Х674ОА124",
                                   "НИВА Chevrolet  Г/н  Е594ЕТ19",
                                   "Fiat Fullback Г/н У746ТЕ124",
                                   "JAC Т6  Г/н  В808ТМ124",
                                   "HINO Ranger Г/н В195КУ19",
                                   "Mitsubishi L200 Г/н А642РУ124",
                                   "ГАЗ Г/н У708УР124",
                                   "GREAT WALL POER Г/н Т891РК124",
                                   "JAC Т6  Г/н  К515ТМ124",
                                   "НИВА Г/н К709РН1244",
                                   "КАМАЗ Г/н О528ХА124",
                                   "Mitsubishi L200 Г/н К961МТ124",
                                   "МАЗ Г/н А233РЕ124",
                                   "HINO Ranger Г/н Х074КМ19",
                                   "Mitsubishi L200 Г/н К863ТВ124",
                                   "НИВА Г/н Е502ОР124",
                                   "КАМАЗ Г/н М473АК124"])

        #Заполняем водителей
        self.CBDriver.addItems(["Каменев И. В.",
                                "Рыжков С. А.",
                                "Метелкин Д. С.",
                                "Жигунов В. А.",
                                "Агапов В. А.",
                                "Арбузин И. В.",
                                "Газизов И. О.",
                                "Коробейников А. В.",
                                "Краско А. В.",
                                "Марцыновский Р. В.",
                                "Серединский М. И.",
                                "Хуторской Д. Н.",
                                "Газизов А. В.",
                                "Большаков В. С.",
                                "Лебеденко Ю. В.",
                                "Петров И. С.",
                                "Голынский В. Н.",
                                "Евдокимов В. В.",
                                "Лященко С. Н.",
                                "Татти Л. И.",
                                "Степаненко А. В.",
                                "Калашников С. С.",
                                "Даровских В. С.",
                                "Красильников А. Ю.",
                                "Степень В.В.",
                                "Щенников В. А. ",
                                "Елисейкин А. Н.",
                                "Болотов Д.Н.",
                                "Чебодаев В. А."])

        #Сортируем транспорт и водителей
        self.CBTransport.model().sort(0)
        self.CBDriver.model().sort(0)

        #Заполняем еденицу измерения ТМЦ
        for cb_piece in self.findChildren(QComboBox, QRegExp("CBPiece\d+")):
            cb_piece.addItems(["Шт.", "Л.", "М2"])

        #Заполняем номер накладной исходя из того есть ли сохраненная накладная
        # либо если ее нет смотрим последнию напечатаную и прибавляем 1
        save_folder = "./Save"
        print_folder = "./Print"
        max_number = 0
        if os.path.exists(save_folder) and os.path.isdir(save_folder):
            for file in os.listdir(save_folder):
                if file.endswith(".xlsx"):
                    file_number = int(''.join(filter(str.isdigit, file)))
                    max_number = max(max_number, file_number)
        if max_number == 0:
            if os.path.exists(print_folder) and os.path.isdir(print_folder):
                for file in os.listdir(print_folder):
                    if file.endswith(".xlsx"):
                        file_number = int(''.join(filter(str.isdigit, file)))
                        max_number = max(max_number, file_number)
            max_number += 1
        self.SBNumber.setValue(max_number)

        #Проверяем есть ли сохраненная накладная и если да то загружаем ее
        self.check_and_fill_form_from_excel()


    #Очистка содержимого формы
    def clearLineEditContents(self):
        for i in range(1, 21):
            le_name = getattr(self, f"LEName{i}", None)
            le_quantity = getattr(self, f"LEQuantity{i}", None)
            if le_name:
                le_name.clear()
            if le_quantity:
                le_quantity.clear()

    def check_and_fill_form_from_excel(self):
        folder_path = "Save"
        files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
        if len(files) == 1:
            file_path = os.path.join(folder_path, files[0])
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            self.CBRecipient.setCurrentText(sheet["C5"].value)
            for i in range(1, 21):
                value = sheet[f"B{9 + i}"].value
                if value is not None:
                    getattr(self, f"LEName{i}").setText(str(value))
                else:
                    getattr(self, f"LEName{i}").clear()
            for i in range(1, 21):
                value = sheet[f"I{9 + i}"].value
                if value is not None:
                    getattr(self, f"LEQuantity{i}").setText(str(value))
                else:
                    getattr(self, f"LEQuantity{i}").clear()
            #for i in range(1, 21):
            #    value = sheet[f"H{9 + i}"].value
            #    if value is not None:
            #        cb_piece = getattr(self, f"CBPiece{i}")
            #        index = cb_piece.findText(str(value))
            #        if index != -1:
            #            cb_piece.setCurrentIndex(index)
            #        else:
            #            cb_piece.addItem(str(value))
            #    else:
            #        cb_piece.clear()
            wb.close()

    def closeEvent(self, event):
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
