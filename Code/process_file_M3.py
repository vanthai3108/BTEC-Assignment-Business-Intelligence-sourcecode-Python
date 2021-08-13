import openpyxl
from colorama import Fore, Back, Style
import emoji

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
    max_row = len(list_data);
    if max_row >= row:
        max_row = row
    for i in range(0, max_row):
        for j in range(0, len(list_data[i])):
            print(str(list_data[i][j])[0:count].ljust(count), end='| ')
        if i == 0:
            print()
            new_count = count
            for j in range(0, len(list_data[i])):
                underline = '_'
                print(underline.ljust(new_count, '_'), end='|')
                if new_count == count:
                    new_count = new_count + 1;
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
            list_row_character.insert(0, i + 1)
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

print()
i = 0;
while(i < 12):
    i=i+1
    print(Style.BRIGHT + Fore.LIGHTBLUE_EX+ emoji.emojize(':deciduous_tree:'), end='')
print(Fore.LIGHTBLUE_EX +" Support Excel Data Processing Application ", end= '')
i=0;
while(i < 12):
    i=i+1
    print(emoji.emojize(':deciduous_tree:'), end='')
while(True):
    print(Style.RESET_ALL)
    print()
    print(Fore.LIGHTCYAN_EX + "============================== ", end= '')
    print(emoji.emojize(':maple_leaf:') + " Apllication Menu "+ emoji.emojize(':maple_leaf:'), end='')
    print(" ==============================")
    print(Fore.LIGHTCYAN_EX + emoji.emojize('                   :backhand_index_pointing_right:')+  " Enter 1:" + Fore.LIGHTWHITE_EX+ " Import data")
    print(Fore.LIGHTCYAN_EX + emoji.emojize('                   :backhand_index_pointing_right:')+  " Enter 2:" + Fore.LIGHTWHITE_EX + " Show data")
    print(Fore.LIGHTCYAN_EX + emoji.emojize('                   :backhand_index_pointing_right:')+  " Enter 3:" + Fore.LIGHTWHITE_EX + " Clear empty data")
    print(Fore.LIGHTCYAN_EX + emoji.emojize('                   :backhand_index_pointing_right:')+  " Enter 4:" + Fore.LIGHTWHITE_EX + " Clear special character")
    print(Fore.LIGHTCYAN_EX + emoji.emojize('                   :backhand_index_pointing_right:')+  " Enter 5:" + Fore.LIGHTWHITE_EX + " Clear special character by column")
    print(Fore.LIGHTCYAN_EX + emoji.emojize('                   :backhand_index_pointing_right:')+  " Enter 6:" + Fore.LIGHTRED_EX + " Exit application")
    print(Fore.LIGHTCYAN_EX + "===================================================================================")
    print(Fore.LIGHTCYAN_EX+emoji.emojize(':backhand_index_pointing_right:') + " Please enter your choice: ", end= '')
    choose = str(input())
    if choose == '1':
        print(Fore.LIGHTWHITE_EX + "------------------------------------Import data------------------------------------")
        print(Fore.LIGHTYELLOW_EX + emoji.emojize(':prohibited:') +  " Note: The application only accepts files with the extension .xlsx ")
        try:
            load_data(wb, path)
            print( Fore.LIGHTYELLOW_EX + emoji.emojize(':prohibited:') + " Already imported data, do you want to import new data? ")
            print("   Enter 1: Import new data")
            print("   Enter 2: Back to menu")
            check_choose = True
            while(check_choose):
                print( Fore.LIGHTWHITE_EX + emoji.emojize('   :backhand_index_pointing_right:')+ " Please enter your choice:", end= ' ')
                choose_new = str(input())
                if choose_new == '1':
                    print( Fore.LIGHTWHITE_EX + emoji.emojize('   :backhand_index_pointing_right:')+ " Please enter the path to your file: ", end=' ')
                    path = str(input())
                    try:
                        wb, sheet, data_list = load_new_data(path)
                        print(Fore.LIGHTGREEN_EX + emoji.emojize( '   :check_mark:') + " Import data successfull")
                        check_choose = False
                    except:
                        print(Fore.LIGHTRED_EX+ emoji.emojize( '   :cross_mark:') + " The path is not correct !")
                elif choose_new == '2':
                    print(Fore.LIGHTGREEN_EX+ "   You have selected back to the menu!")
                    check_choose = False
                else:
                    print(Fore.LIGHTRED_EX + emoji.emojize( '   :cross_mark:') + " Incorrect selection, please re-enter ")
        except:
            print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:')+ " Please enter the path to your file: ", end=' ')
            path = str(input())
            try:
                wb, sheet, data_list = load_new_data(path)
                print(Fore.LIGHTGREEN_EX + emoji.emojize( ':check_mark:') + " Import data successful!")
            except:
                print(Fore.LIGHTRED_EX+ emoji.emojize( ':cross_mark:') + " The path is not correct !")

    elif choose == '2':
        print(Fore.LIGHTWHITE_EX + "-------------------------------------Show data-------------------------------------")
        try:
            print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Enter the width of the columns (characters): ", end='')
            count = int(input())
            print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Enter the maximum number of rows displayed: ", end='')
            row = int(input())
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " Results:")
            print(Fore.LIGHTCYAN_EX)
            show_data(sheet, count, row)
        except:
            print(Fore.LIGHTRED_EX+  emoji.emojize(':cross_mark:') + " No data to display")

    elif choose == '3':
        print(Fore.LIGHTWHITE_EX + "----------------------------------Clear empty data---------------------------------")
        try:
            characters = ['None']
            print(Fore.LIGHTGREEN_EX + "...Data is being processed, please wait a moment!")
            clear_data(data_list, characters)
            characters = []
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " Done processing!")
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " All data lines contain empty data have been successfully removed!")
            data_list = load_data(wb, path)
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " Data has been loaded!")
        except:
            print(Fore.LIGHTRED_EX + emoji.emojize(':cross_mark:') + " No data to process or error")
    elif choose == '4':
        print(Fore.LIGHTWHITE_EX + "------------------------------Clear special character------------------------------")
        try:
            print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Please enter the characters to remove, type 'exit' to end.")
            character = ""
            characters = []
            while character != "exit":
                print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Enter the character to be removed:  ", end='')
                character = str(input())
                if character == "":
                    print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Invalid character, please re-enter:  ", end='')
                elif character == "exit":
                    print(Fore.LIGHTGREEN_EX + "...Data is being processed, please wait a moment!")
                else:
                    characters.append(character)
            clear_data(data_list, characters)
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " Done processing!")
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " All data lines contain character " + Fore.LIGHTWHITE_EX, end='')
            for character in characters:
                print("'"+character+"', ",end='')
            print(Fore.LIGHTGREEN_EX + " have been successfully removed!")
            characters = []
            data_list = load_data(wb, path)
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " Data has been loaded!")
        except:
            print(Fore.LIGHTRED_EX + emoji.emojize(':cross_mark:') + " No data to process or error")
    elif choose == '5':
        print(Fore.LIGHTWHITE_EX + "------------------------Clear special character by column--------------------------")
        try:
            print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Enter the column name to be cleaned : ", end='')
            column = str(input())
            index_column = -1;
            for j in range(0, len(data_list[0])):
                if column == str(data_list[0][j]):
                    index_column = j
                    break
            if index_column == -1:
                print(Fore.LIGHTRED_EX + emoji.emojize(':cross_mark:') + " Column name does not exist!")
                continue
            else:
                print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Please enter the characters to remove, type 'exit' to end.")
                character = ""
                characters = []
                while character != "exit":
                    print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Enter the character to be removed:  ", end='')
                    character = str(input())
                    if character == "":
                        print(Fore.LIGHTWHITE_EX + emoji.emojize(':backhand_index_pointing_right:') + " Invalid character!", end='')
                    elif character == "exit":
                        print(Fore.LIGHTGREEN_EX + "...Data is being processed, please wait a moment!")
                    else:
                        characters.append(character)
            clear_data_row(data_list, characters, index_column)
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " Done processing!")
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " All data lines contain special character of column '" + Fore.LIGHTWHITE_EX + column + Fore.LIGHTGREEN_EX + "' have been removed!")
            characters = []
            data_list = load_data(wb, path)
            print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " Data has been loaded!")
        except:
            print(Fore.LIGHTRED_EX+ emoji.emojize(':cross_mark:') + " No data to process or error")
    elif choose == '6':
        print(Fore.LIGHTWHITE_EX + "----------------------------------Exit application---------------------------------")
        print(Fore.LIGHTGREEN_EX + emoji.emojize(':check_mark:') + " You have chosen to exit the application")
        print(Fore.LIGHTGREEN_EX + emoji.emojize(':OK_hand:') + " Good bye and see you later")
        exit()
    else:
        print(Fore.LIGHTRED_EX + emoji.emojize(':cross_mark:') + " Incorrect selection!")