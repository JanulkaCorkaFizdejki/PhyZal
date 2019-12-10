import sys
import xlrd
loc = ("oceny.xlsx")
wb = xlrd.open_workbook(loc)
count_sheet = wb.nsheets


def sort_dict (keys = [], values = []):
    dict = {}
    counter = 0
    for item in keys:
        dict.update({item: values[counter]})
        counter += 1
    for k,v in sorted(dict.items(), key=lambda x:x[1], reverse=True):
        print("{} - {}".format(k, v))


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


# Ranking ocen z poszczególnych przedmiotów
def rank1 ():
    print("\nRANIKING OCEN Z POSZCZEGÓLNYCH PRZEDNIOTÓW:")
    sheet_names = wb.sheet_names()
    for index in range(0, count_sheet):
        curr_sheet = wb.sheet_by_index(index)
        names = []
        eval = []
        for row_index in range(0, curr_sheet.nrows):
            names.append(curr_sheet.cell(row_index, 0).value)
            eval.append(curr_sheet.cell(row_index, 1).value)
        print(sheet_names[index])
        print("___________________")
        sort_dict(names, eval)

# Ranking średnich ocen z wszystkich przedmiotów
def rankAll ():
    print("\nRANIKING ŚREDNICH OCEN Z WSZYSTKICH PRZEDNIOTÓW:")
    eval = []
    names = []
    for index in range(0, count_sheet):
        curr_sheet = wb.sheet_by_index(index)
        counter = 0
        for row_index in range(0, curr_sheet.nrows):
            if index == 0:
                eval.append(curr_sheet.cell(row_index, 1).value)
            elif index > 0 and index < count_sheet - 1:
                eval[counter] = eval[counter] + curr_sheet.cell(row_index, 1).value
            else:
                eval[counter] = round((eval[counter] + curr_sheet.cell(row_index, 1).value) / count_sheet, 2)
                names.append(curr_sheet.cell(row_index, 0).value)
            counter += 1
    sort_dict(names, eval)

print("\nSTART")
print("---------------------------------")
print("Lista przedmiotów: Matematyka, Informatyka, Język angileski")
print("\nANALIZA WYNIKÓW NAUCZANIA")
i = input("Czy chcesz poznać średnią ocen dla danego przedmiotu?[y/n] ")
if (i == "y"):
    p = input("Podaj nazwę przedmiotu: ")
    pp = avg_one(p)
    if pp > 1:
        print("Przedmiot [{}] : {}".format(p, avg_one(p)))
    else:
        print("Studenci nie uczyli się tego przedmiotu!")

i = input("Czy chcesz poznać średnią ocen danego studenta z wszystkich przedmiotów?[y/n] ")
if (i == "y"):
    n = input("Podaj imię i nazwisko studenta: ")
    s = avg_student(n)
    if s > 1:
        print("Średnia ocen dla [{}] wynosi: {}".format(n, s))
    else:
        print("Taki student nie istnieje!")

i = input("Czy chcesz poznać średnią ocen wszystkich studentów z wszystkich przedmiotów?[y/n] ")
if (i == "y"):
    print("Średnia ocen wszystkich studentów z wszystkich przedmiotów wynosi: {}".format(avg_students_all()))

i = input("Czy chcesz poznać ranking najlepszych studnetów pod wzglądem ocen z poszczególnych przedmiotów?[y/n] ")
if (i == "y"):
    rank1()

i = input("Czy chcesz poznać ranking najlepszych studnetów pod wzglądem średniaj ocen z wszystkich przedmiotów?[y/n] ")
if (i == "y"):
    rankAll()

print("*********************************************")
print("KONIEC")
print("*********************************************")









