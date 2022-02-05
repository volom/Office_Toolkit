# the script helps you to save info from clickboard and then update it with accumulated info
import pyperclip

# list with accumulated info
acc_info = []

# sep for info
sep = ','

while True:
    pyperclip.waitForNewPaste()
    cb_value = pyperclip.paste()
    acc_info.append(cb_value)
    cb_value_acc = f'{sep} '.join(acc_info)
    pyperclip.copy(cb_value_acc)
    print(cb_value_acc)