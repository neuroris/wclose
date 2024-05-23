from pywinauto.application import Application
import pywinauto
import pyautogui
import time, os, sys
import re

class WookClose:
    def __init__(self):
        self.failed = False
        self.fail_message = '--------- Failed app list ---------'
        self.close_process()

    def close_process(self):
        print('Wook close-up process started.\n')

        # Task bar application
        taskbar_app = Application(backend='uia')
        taskbar_app.connect(class_name='Shell_TrayWnd')
        self.taskbar_dlg = taskbar_app.window(class_name='Shell_TrayWnd')
        # self.taskbar_dlg.print_control_identifiers()

        # Close applications
        # self.close_ADT()
        # self.close_hanaro()
        # self.close_pointnix()
        # self.close_excel()
        # self.close_explorer()
        # self.close_chrome()
        # self.close_pycharm()
        self.close_app('하나로')

        if self.failed:
            print('\nSome app failed to close.\n')
            print(self.fail_message)
        else:
            print('\nWhole close-up process done successfully!')

    def report_failure(self, message, e=None):
        self.failed = True
        self.fail_message += '\n' + message

        if e is not None:
            print('========= Exception Occur =========')
            print(e)
            print('===================================')

    def get_top_dlg(self, window_title):
        maximum_trial = 20
        waiting_time = 1

        app = Application('uia')
        for count in range(1, maximum_trial):
            try:
                app.connect(title_re=window_title)
                print('{} app is now connected.'.format(window_title))
                break
            except:
                print('{} app is not ready. Waiting {}s...trial({})'.format(window_title, waiting_time, count))
                time.sleep(waiting_time)

        dlg = app.window(title_re=window_title)
        for count in range(1, maximum_trial):
            if dlg.exists():
                print('{} dlg is now active.'.format(window_title))
                return dlg
            else:
                print('{} dlg is not active. Waiting {}s...trial({})'.format(window_title, waiting_time, count))
                time.sleep(waiting_time)

        print('Getting {} dlg failed!!!'.format(window_title))
        self.report_failure(window_title)

        return None

    def close_app(self, title, app_name=None):
        try:
            app_name = app_name if app_name else title
            print(title + ' close up procedure')
            taskbar_dlg = self.taskbar_dlg.window(title_re=title + '.*개의 실행 중인 창.*')
            if not taskbar_dlg.exists():
                print(title + ' is not running')
                return

            taskbar_dlg.right_click_input()
            pywinauto.keyboard.send_keys('{UP}')
            pywinauto.keyboard.send_keys('{ENTER}')
            pywinauto.keyboard.send_keys('{ENTER}')
        except Exception as e:
            self.report_failure(title, e)

    def close_hanaro(self):
        try:
            print('Hanaro close up procedure')
            taskbar_hanaro_dlg = self.taskbar_dlg.window(title_re='하나로 .* 1개의 실행 중인 창.*')

            if not taskbar_hanaro_dlg.exists():
                print('Hanaro is not running')
                return

            # taskbar_hanaro_dlg.click_input()
            # hanaro_dlg = self.get_top_dlg('하나로')
            # hanaro_dlg.close()
            # pywinauto.keyboard.send_keys('{ENTER}')
            taskbar_hanaro_dlg.right_click_input()
            pywinauto.keyboard.send_keys('{UP}')
            pywinauto.keyboard.send_keys('{ENTER}')
            pywinauto.keyboard.send_keys('{ENTER}')
        except Exception as e:
            self.report_failure('Hanaro', e)

    def close_pointnix(self):
        try:
            print('Pointnix close up procedure')
            taskbar_pointnix_dlg = self.taskbar_dlg.window(title_re='^CDX-View .*개의 실행 중인 창')
            if not taskbar_pointnix_dlg.exists():
                print('Pointnix is not running')
                return

            taskbar_pointnix_dlg.right_click_input()
            time.sleep(0.5)
            pywinauto.keyboard.send_keys('{UP}')
            pywinauto.keyboard.send_keys('{ENTER}')
        except Exception as e:
            self.report_failure('PointNix', e)

    def close_ADT(self):
        try:
            print('ADT close up procedure')
            pyautogui.hotkey('ctrl','win','left')
            pyautogui.hotkey('ctrl','win','left')
            taskbar_ADT_dlg = self.taskbar_dlg.window(title_re='ADT.* 1개의.*')
            if not taskbar_ADT_dlg.exists():
                print('ADT is not running')
                return

            # adt_dlg = self.get_top_dlg('ADT')
            # adt_close_button = adt_dlg['Button6']
            # adt_close_button.click_input()
            taskbar_ADT_dlg.right_click_input()
            pywinauto.keyboard.send_keys('{UP}')
            pywinauto.keyboard.send_keys('{ENTER}')
            pyautogui.hotkey('ctrl','win','right')
        except Exception as e:
            self.report_failure('ADT', e)

    def close_excel(self):
        try:
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
            app.connect(path='C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE')

            for i in range(num_window):
                taskbar_excel_dlg.click_input()
                pywinauto.keyboard.send_keys('{RIGHT}')
                pywinauto.keyboard.send_keys('{ENTER}')
                excel_dlg = app.top_window()
                excel_save_dlg = excel_dlg['저장']
                excel_save_dlg.click_input()
                excel_close_dlg = excel_dlg['닫기']
                excel_close_dlg.click_input()

            # for i in range(num_window):
            #     taskbar_excel_dlg.click_input()
            #     pywinauto.keyboard.send_keys('{UP}')
            #     pywinauto.keyboard.send_keys('{ENTER}')
            #     excel_dlg = app.top_window()
            #     excel_save_dlg = excel_dlg.window(title='저장', top_level_only=True, control_type='Button')
            #     excel_save_dlg.click_input()
            #     excel_dlg.close()
        except Exception as e:
            self.report_failure('Excel', e)

    def close_explorer(self):
        try:
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
        except Exception as e:
            self.report_failure('Explorer', e)

    def close_chrome(self):
        try:
            print('Chrome close up procedure')
            # taskbar_chrome_dlg = self.taskbar_dlg.window(title_re='Chrome .* 실행 중인 창')
            taskbar_chrome_dlg = self.taskbar_dlg.window(title_re='Google Chrome .* 실행 중인 창')
            if not taskbar_chrome_dlg.exists():
                print('No Chrome browser is running')
                return

            taskbar_chrome_dlg.right_click_input()
            time.sleep(0.5)
            pywinauto.keyboard.send_keys('{UP}')
            pywinauto.keyboard.send_keys('{ENTER}')
        except Exception as e:
            self.report_failure('Chrome', e)

    def close_pycharm_deprecated(self):
        try:
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
        except Exception as e:
            self.report_failure('PyCharm', e)

    def close_pycharm(self):
        try:
            print('Pycharm close up procedure')
            taskbar_pycharm_dlg = self.taskbar_dlg.window(title_re='PyCharm.* 실행 중인 창')
            if not taskbar_pycharm_dlg.exists():
                print('PyCharm is not running.')
                return

            text = taskbar_pycharm_dlg.texts()
            title = text[0]
            num_window_str = re.findall(r'\d+', title)
            num_window = int(num_window_str[-1])

            taskbar_pycharm_dlg.right_click_input()
            time.sleep(0.5)
            pywinauto.keyboard.send_keys('{UP}')
            pywinauto.keyboard.send_keys('{ENTER}')

            if num_window > 1:
                pycharm_close_dlg = self.get_top_dlg('Welcome to PyCharm')
                pycharm_close_dlg.close()
        except Exception as e:
            self.report_failure('PyCharm', e)

if __name__ == '__main__':
    wc = WookClose()