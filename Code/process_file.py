import openpyxl

def load_new_data(path):
    wb = openpyxl.load_workbook(path)
    sheet = wb[wb.sheetnames[0]]
    data_list = list(sheet.values)
    return wb, sheet, data_list

def load_data(wb, path):
    wb.save(path)
    wb, sheet, data_list = load_new_data(path)
    return data_list

def show_data(sheet_data, count, row):
    list_data = list(sheet_data.values)
    max_row = len(list_data)
    if max_row >= row:
        max_row = row
    for i in range(0, max_row):
        for j in range(0, len(list_data[i])):
            print(str(list_data[i][j])[0:count].ljust(count), end=' | ')
        print()
    print()

def delete_rows(list_del, sheet):
    for x in list_del:
        sheet.delete_rows(idx=x)

def clear_data(list_data, characters):
    list_row_character = []
    for i in range(1, len(list_data)):
        list_data[i] = list(list_data[i])
        check = True
        for j in range(0, len(list_data[i])):
            for character in characters:
                if character in str(list_data[i][j]):
                    check = False
                    break
        if check == False:
            list_row_character.insert(0, i+1)
    delete_rows(list_row_character, sheet)

def clear_data_row(list_data, characters, index_column):
    list_row_character = []
    check = True
    for i in range(1, len(list_data)):
        list_data[i] = list(list_data[i])
        for character in characters:
            if character in str(list_data[i][index_column]):
                check = False
        if check == False:
            list_row_character.insert(0, i+1)
            check = True
    delete_rows(list_row_character, sheet)

while(True):
    print()
    print("================================Menu=================================")
    print("Enter 1: Import data.")
    print("Enter 2: Show data.")
    print("Enter 3: Clear empty data.")
    print("Enter 4: Clear special character.")
    print("Enter 5: Clear special character by column.")
    print("Enter 6: Exit application.")
    print("=====================================================================")
    print("Enter your choice: ", end= '')
    choose = str(input())
    if choose == '1':
        print("-----------------------------Import data-----------------------------")
        print("Please enter the path to your file: ", end=' ')
        path = str(input())
        try:
            wb, sheet, data_list = load_new_data(path)
            print("Import data successful!")
        except:
            print("The path is not correct!")
    elif choose == '2':
        print("------------------------------Show data------------------------------")
        try:
            print("Enter the width of the columns (characters): ", end='')
            count = int(input())
            print("Enter the maximum number of rows displayed: ", end='')
            row = int(input())
            print("Results:")
            show_data(sheet, count, row)
        except:
            print("No data to display")
    elif choose == '3':
        print("---------------------------Clear empty data--------------------------")
        try:
            characters = ['None']
            clear_data(data_list, characters)
            print("All data lines contain empty data have been successfully removed!")
            data_list = load_data(wb, path)
            print("Data has been loaded!")
        except:
            print("No data to process")
    elif choose == '4':
        print("-----------------------Clear special character-----------------------")
        try:
            print("Please enter the characters to remove, type 'exit' to end.")
            character = ""
            characters = []
            while character != "exit":
                print("Enter the character to be removed:  ", end='')
                character = str(input())
                if character == "":
                    print("Invalid character, please re-enter:  ", end='')
                elif character == "exit":
                    print("")
                else:
                    characters.append(character)
            clear_data(data_list, characters)
            print("All data lines contain spacial character have been successfully removed!")
            data_list = load_data(wb, path)
            print("Data has been loaded!")
        except:
            print("No data to process")
    elif choose == '5':
        print("-----------------Clear special character by column-------------------")
        try:
            print("Enter the column name to be cleaned : ", end='')
            column = str(input())
            index_column = -1;
            for j in range(0, len(data_list[0])):
                if column == str(data_list[0][j]):
                    index_column = j
                    break
            if index_column == -1:
                print("Column name does not exist!")
            else:
                print("Please enter the characters to remove, type 'exit' to end.")
                character = ""
                characters = []
                while character != "exit":
                    print("Enter the character to be removed:  ", end='')
                    character = str(input())
                    if character == "":
                        print("Invalid character!", end='')
                    elif character == "exit":
                        print("")
                    else:
                        characters.append(character)
            clear_data_row(data_list, characters, index_column)
            print("All data lines contain spacial character ", end='')
            print(" of column '" + column + "' have been successfully removed!")
            data_list = load_data(wb, path)
            print("Data has been loaded!")
        except:
            print("No data to process")
    elif choose == '6':
        print("---------------------------Exit application--------------------------")
        print("You have chosen to exit the application")
        exit()
    else:
        print("Incorrect selection!")