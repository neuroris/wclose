from pywinauto.application import Application
import pywinauto
import time, os, sys
import re

class WookClose:
    def __init__(self):
        self.close_process()

    def close_process(self):
        print('Wook close-up process started.\n')

        # Task bar application
        taskbar_app = Application(backend='uia')
        taskbar_app.connect(class_name='Shell_TrayWnd')
        self.taskbar_dlg = taskbar_app.window(class_name='Shell_TrayWnd')

        # Close applications
        try:
            self.close_hanaro()
            self.close_pointnix()
            self.close_ADT()
            self.close_excel()
            self.close_explorer()
            self.close_chrome()
            self.close_pycharm()
        except Exception as e:
            print(e)

        print('Whole close-up process done successfully!\n')

    def close_hanaro(self):
        print('Hanaro close up procedure')
        taskbar_hanaro_dlg = self.taskbar_dlg.window(title_re='하나로 .* 1개의 실행 중인 창.*')
        if not taskbar_hanaro_dlg.exists():
            print('Hanaro is not running')
            return

        taskbar_hanaro_dlg.click_input()
        app = Application('uia')
        app.connect(title_re='^하나로 OK.*')
        hanaro_dlg = app.window(title_re='^하나로 OK.*')
        hanaro_dlg.close()
        pywinauto.keyboard.send_keys('{ENTER}')

    def close_pointnix(self):
        print('Pointnix close up procedure')
        taskbar_pointnix_dlg = self.taskbar_dlg.window(title_re='^CDX-View .*개의 실행 중인 창')
        if not taskbar_pointnix_dlg.exists():
            print('Pointnix is not running')
            return

        taskbar_pointnix_dlg.right_click_input()
        time.sleep(0.5)
        pywinauto.keyboard.send_keys('{UP}')
        pywinauto.keyboard.send_keys('{ENTER}')

    def close_ADT(self):
        print('ADT close up procedure')
        taskbar_ADT_dlg = self.taskbar_dlg.window(title_re='ADT EYE .* 1개의.*')
        if not taskbar_ADT_dlg.exists():
            print('ADT is not running')
            return

        app = Application('uia')
        app.connect(title_re='ADT EYE 2.0')
        adt_dlg = app.window(title_re='ADT EYE 2.0')
        adt_dlg.close()

    def close_excel(self):
        print('Excel close up procedure')
        taskbar_excel_dlg = self.taskbar_dlg.window(title_re='Excel.* 실행 중인 창')
        if not taskbar_excel_dlg.exists():
            print('Excel is not running.')
            return

        text = taskbar_excel_dlg.texts()
        title = text[0]
        num_window_str = re.findall(r'\d+', title)
        num_window = int(num_window_str[0])

        app = Application('uia')
        # app.connect(path='C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXE')
        app.connect(path='C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE')
        for i in range(num_window):
            taskbar_excel_dlg.click_input()
            pywinauto.keyboard.send_keys('{UP}')
            pywinauto.keyboard.send_keys('{ENTER}')
            excel_dlg = app.top_window()
            excel_save_dlg = excel_dlg.window(title='저장', top_level_only=True, control_type='Button')
            excel_save_dlg.click_input()
            excel_dlg.close()

    def close_explorer(self):
        print('Explorer close up procedure')
        taskbar_explorer_dlg = self.taskbar_dlg.window(title_re='파일 탐색기 .* 실행 중인 창')
        if not taskbar_explorer_dlg.exists():
            print('No explorer is running')
            return

        # self.app.connect(class_name='CabinetWClass')
        taskbar_explorer_dlg.right_click_input()
        time.sleep(0.5)
        pywinauto.keyboard.send_keys('{UP}')
        pywinauto.keyboard.send_keys('{ENTER}')

    def close_chrome(self):
        print('Chrome close up procedure')
        taskbar_chrome_dlg = self.taskbar_dlg.window(title_re='Chrome .* 실행 중인 창')
        if not taskbar_chrome_dlg.exists():
            print('No Chrome browser is running')
            return

        taskbar_chrome_dlg.right_click_input()
        time.sleep(0.5)
        pywinauto.keyboard.send_keys('{UP}')
        pywinauto.keyboard.send_keys('{ENTER}')

    def close_pycharm(self):
        print('Pycharm close up procedure')
        taskbar_pycharm_dlg = self.taskbar_dlg.window(title_re='PyCharm.* 실행 중인 창')
        if not taskbar_pycharm_dlg.exists():
            print('PyCharm is not running.')
            return

        text = taskbar_pycharm_dlg.texts()
        title = text[0]
        num_window_str = re.findall(r'\d+', title)
        num_window = int(num_window_str[-1])

        for i in range(num_window):
            taskbar_pycharm_dlg.right_click_input()
            time.sleep(0.5)
            pywinauto.keyboard.send_keys('{UP}')
            pywinauto.keyboard.send_keys('{ENTER}')

if __name__ == '__main__':
    wc = WookClose()