###########################################
#       Экспорт ФИО из БД ЭЖД в Moodle    #
#               version 0.1               #
###########################################


import openpyxl
import transliterate


#TODO: Открытие файла Excel
nameFileExcel = 'test.xlsx'
wb = openpyxl.load_workbook(nameFileExcel)      #открытие книги Excel
listSheets = wb.sheetnames                      #получение списка листов книги Excel
sheets1 = wb[listSheets[0]]                     #берем первый лист книги Excel
# print(sheets1.cell(row=3,column=2).value)     #получение содержимого ячейки

def printInfoToConsole(rusLastName, rusFirstName, lastName, e_mail):
    """печать в консоль данных"""
    print(rusLastName + ' ' + rusFirstName + '; login: ' +
          lastName + '; password: ' + lastName +
          '; e-mail: ' + e_mail)

def translitNames(fio):
    """получение транслита фамилии и имени для логина, пароля и e-mail"""
    lastName = transliterate.translit(fio[0], reversed=True)
    lastName = lastName.replace("'", "")
    lastName = lastName.swapcase()[0] + lastName[1:]
    firstName = transliterate.translit(fio[1], reversed=True)
    firstName = firstName.replace("'", "")
    firstName = firstName.swapcase()[0] + firstName[1:]
    return lastName, firstName

def e_mailGet(lastName, firstName):
    e_mail = lastName + '.' + firstName + '@test.ru'
    return e_mail

#TODO: Записать Ф и И в новый лист Excel и сохранить в csv

kItem = 0                                       #кол-во записей

for i in range(3,51):
    # TODO: Получение столбца с ФИО
    valueCurrentCell = sheets1.cell(row=i,column=2).value
    if str(type(valueCurrentCell)) != "<class 'NoneType'>":
        try:
            fio = valueCurrentCell.split()
            # TODO: Разбить ФИО на фамилию и имя, удалить отчество
            rusLastName = fio[0]
            rusFirstName = fio[1]
            lastName = translitNames(fio)[0]
            firstName = translitNames(fio)[1]
            e_mail = e_mailGet(lastName, firstName)
            printInfoToConsole(rusLastName, rusFirstName, lastName, e_mail)
            if e_mail:
                kItem = kItem + 1
        except:
            pass
print('Total: ' + str(kItem))
