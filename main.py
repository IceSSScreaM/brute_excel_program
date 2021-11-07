import itertools
from string import digits, punctuation, ascii_letters
import win32com.client as client
import time
import os


def ask_long():
    while True:
        while True:
            try:
                minLong = int(input("Введите наименьшее количество символов: "))
                break
            except:
                os.system("cls")
                print("[!] Проверьте введенные данные [!]\n")


        os.system("cls")
        while True:
            try:
                maxLong = int(input("Введите наибольшее количество символов: "))
                break
            except:
                os.system("cls")
                print("[!] Проверьте введенные данные [!]\n")

        if (minLong > maxLong):
            os.system("cls")
            print("[!] Введенные значения некорректны [!]\n")
        else:
            break


    return minLong, maxLong


def ask_possibleSimbols():
    os.system("cls")
    possibleSymbols = ""
    
    os.system("cls")
    while True:
        print("Использовали ли вы цифры в пароле? (+ или -): ")
        answer = input()
        if (answer == "+"):
            possibleSymbols = possibleSymbols + digits
            break
        elif (answer == "-"):
            break
        else:
            os.system("cls")
            print("[!] Проверьте введенные данные [!]\n")

    os.system("cls")
    while True:
        print("Использовали ли вы латиницу в пароле? (+ или -): ")
        answer = input()
        if (answer == "+"):
            possibleSymbols = possibleSymbols + ascii_letters
            break
        elif (answer == "-"):
            break
        else:
            os.system("cls")
            print("[!] Проверьте введенные данные [!]\n")

    os.system("cls")
    while True:
        print("Использовали ли вы специальные в пароле? (+ или -): ")
        answer = input()
        if (answer == "+"):
            possibleSymbols = possibleSymbols + punctuation
            break
        elif (answer == "-"):
            break
        else:
            os.system("cls")
            print("[!] Проверьте введенные данные [!]\n")

    return possibleSymbols

    
def ask_path():
    os.system("cls")
    print("Укажите путь к файлу: ")
    path = input()

    return path


def brute_excel_doc(minLong, maxLong, possibleSymbols, path):

    count = 0
    for pass_length in range(minLong, maxLong+1):
        for password in itertools.product(possibleSymbols, repeat=pass_length):
        
            password = "".join(password)

            opened_doc = client.Dispatch("Excel.Application")
            count += 1

            if (count % 100 == 0): 
                os.system("cls")

            try:
                opened_doc.Workbooks.Open(r"" + path, False, True, None, password)

                time.sleep(0.1)

                return f"Верный пароль: {password}"
            except:
                print(f"Попытка #{count}: {password}")
                pass


def show_result(result):
    os.system("cls")
    if (result == None):
        print("[!] Не удалось подобрать пароль [!]")
        input()
    else:
        print(result)
        input()






def main():
    minLong, maxLong = ask_long()
    possibleSymbols = ask_possibleSimbols()
    path = ask_path()
    result = brute_excel_doc(minLong, maxLong, possibleSymbols, path)
    show_result(result)

    


if __name__ == '__main__':
    main()
