import xlrd
loc = ("oceny.xlsx")
wb = xlrd.open_workbook(loc)
count_sheet = wb.nsheets

# Średnia ocen dla danego przedmiotu
def avg_one (name):
    names = wb.sheet_names()
    if name in names:
        curr_sheet = wb.sheet_by_name(name)
        sum = 0
        counter = 0
        for row_index in range(curr_sheet.nrows):
            cell = curr_sheet.cell(row_index, 1)
            sum += cell.value
            counter += 1
        return round((sum / counter), 2)
    else:
        return 0

# Średnia ocen z wszystkich przedmiotów dla danego studenta
def avg_student (student_name):
    sum = 0
    for index in range(0, count_sheet):
        curr_sheet = wb.sheet_by_index(index)
        for row_index in range(curr_sheet.nrows):
            if student_name == curr_sheet.cell(row_index, 0).value:
                sum += curr_sheet.cell(row_index, 1).value
    return round((sum / count_sheet), 2)

# Średnia ocen wszystkich uczniów dla wszystkich przedmiotów
def avg_students_all ():
    sum = 0
    for index in range(0, count_sheet):
        curr_sheet = wb.sheet_by_index(index)
        rang = curr_sheet.nrows
        for row_index in range(curr_sheet.nrows):
            sum += curr_sheet.cell(row_index, 1).value
    return round((sum / (rang * count_sheet)), 2)

print("\nŚrednia ocen dla danego przedmiotu:")
print("__________________________________")
print("Matematyka: {}".format(avg_one("Matematyka")))
print("Informatyka: {}".format(avg_one("Informatyka")))
print("Język angileski: {}".format(avg_one("Język angielski")))
print("__________________________________\n")

print("\nŚrednia ocen z wszystkich przedmiotów dla danego studenta (wybrane nazwiska):")
print("__________________________________")
print("Albert Nowakowski: {}".format(avg_student("Albert Nowakowski")))
print("Izabela Nowak: {}".format(avg_student("Izabela Nowak")))
print("Onufry Nowak: {}".format(avg_student("Onufry Nowak")))

print("\nŚrednia ocen wszystkich uczniów dla wszystkich przedmiotów:")
print("__________________________________")
print(avg_students_all())



