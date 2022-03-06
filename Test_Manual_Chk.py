import pandas
import time
import pyautogui

excel_data = \
    pandas.read_excel(r'C:\Users\Rutland Products\Documents\Excel Docs for Python Test\AP\Manual Check Template.xlsx',
                      sheet_name=0)
count = 0
time.sleep(4)


for column in range(len(excel_data['Vendor #'])):

    pyautogui.typewrite('M')
    pyautogui.press('enter')
    pyautogui.typewrite(str(int(excel_data['Vendor #'][count])))
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.typewrite(str(int(excel_data['Voucher #'][count])))
    pyautogui.press('Enter')
    pyautogui.press('Enter')
    if count == 0:
        pyautogui.press('f1')  # Change to press
        pyautogui.keyDown('ctrl')  # Change to keydown
        pyautogui.press('c')  # Change to press
        pyautogui.keyUp('ctrl')  # Change to key up
    else:
        if excel_data['Vendor #'][count] != excel_data['Vendor #'][count-1]:
            pyautogui.press('f1')  # Change to press
            pyautogui.keyDown('ctrl')  # Change to keydown
            pyautogui.press('c')  # Change to press
            pyautogui.keyUp('ctrl')  # Change to key up
        else:
            pyautogui.keyDown('ctrl')  # change to key down
            pyautogui.press('v')  # change to press
            pyautogui.keyUp('ctrl')  # change to key up
    pyautogui.press('enter')
    pyautogui.typewrite(str(excel_data['Check Date'][count]).zfill(6))
    pyautogui.press('enter')
    pyautogui.typewrite(str(excel_data['Gross A/P Amt'][count]))
    pyautogui.press('enter')
    pyautogui.typewrite(str(excel_data['Disc Taken'][count]))
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    count += 1
