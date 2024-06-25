import os
import unittest
import warnings
import winshell
from appium import webdriver
from appium.webdriver.common.appiumby import AppiumBy
from appium.options.windows import WindowsOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import socket
from appium.webdriver.appium_service import AppiumService
import subprocess
import pyautogui
import time
import yaml
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, TimeoutException

current_dir = os.path.dirname(os.path.realpath(__file__))
config_file_path = os.path.join(current_dir, 'testdata.yaml')

with open(config_file_path, 'r') as f:
    config = yaml.safe_load(f)

SHORTCUT_PATH = config['SHORTCUT_PATH']
EXE_DIRECTORY = config['EXE_DIRECTORY']
DEFAULT_PORT = config['DEFAULT_PORT']
APP_EXE_PATH = config['APP_EXE_PATH']
APP_NAME = config['APP_NAME']
PASSWORD = config['PASSWORD']

class AppiumServer:
    def __init__(self):
        self.appium_service = AppiumService()
        self.port = DEFAULT_PORT

    def start(self):
        while self.is_port_in_use(self.port):
            self.port += 1
        self.appium_service.start(args=['-p', str(self.port), '--base-path', '/wd/hub'])

    def stop(self):
        self.appium_service.stop()

    @staticmethod
    def is_port_in_use(port):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            return s.connect_ex(('localhost', port)) == 0


class Tests(unittest.TestCase):
    driver = None
    appium_server = AppiumServer()
    target_path = None
    arguments = None

    @classmethod
    def setUpClass(cls):
        global driver, appium_server
        warnings.filterwarnings("ignore")

        Logger.logger_warn('Killing all instances of node.exe to avoid interference of any running Appium sessions')

        os.system('taskkill /f /im node.exe')

        Logger.logger_warn("All instances of 'node' killed, starting Appium server, please wait")

        cls.appium_server.start()

        # Logger.logger_bold('Uninstalling previous version of the app')

        # Functions.uninstall_program(APP_NAME)

        # Parsing shortcut so that if the app needs to be launched with some arguments, they will be parsed automatically and the app will be launched using them
        shortcut = winshell.shortcut(SHORTCUT_PATH)
        cls.target_path = shortcut.path
        cls.arguments = shortcut.arguments


        warnings.filterwarnings("ignore")        

        cls.launch_exe(cls.target_path, cls.arguments, EXE_DIRECTORY)

    @staticmethod
    def launch_exe(target_path, arguments, exe_directory):
        subprocess.Popen([target_path, arguments], cwd=exe_directory)

    def setUp(self) -> None:
        appium_server_url = f"http://127.0.0.1:{self.appium_server.port}/wd/hub"
        options = WindowsOptions()
        options.experimental_webdriver = False
        options.app = "Root"
        options.platform_name = "Windows"
        self.driver = webdriver.Remote(appium_server_url, options=options)

    @classmethod
    def tearDownClass(cls) -> None:
        if cls.driver is not None:
            try:
                cls.driver.quit()
                cls.appium_server.stop()
            except Exception as e:
                Logger.logger_warn(f"An error occurred while stopping the driver and Appium service: {str(e)}")

    '''
    Add new tests here. Any method that starts with 'test_' prefix will be treated as a new test. 
    '''

    def test_update_row(self):
        # Functions.install_program()

        Tests.launch_exe(Tests.target_path, Tests.arguments, EXE_DIRECTORY)

        Functions.open_project(self.driver)

        # Functions.resize_window(800, 600)

        Logger.logger_step(1, "Select a row")
        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "Table View"), 30)

        Functions.context_click_element(self.driver, (AppiumBy.NAME, "Select a row"))

        Logger.logger_step(2, 'Click OK')

        Functions.wait_and_click(self.driver, (AppiumBy.ACCESSIBILITY_ID, "buttonOK"), 20)

        Logger.logger_step(3, 'Right click the row and click update')

        Functions.context_click_element(self.driver, (AppiumBy.ACCESSIBILITY_ID, "row"))

        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "Update"), 20)

        Logger.logger_step(4, 'Enter password')

        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "Download"), 20)

        WebDriverWait(self.driver, 150).until(EC.presence_of_element_located((AppiumBy.ACCESSIBILITY_ID, "textBox"))).send_keys(PASSWORD)

        Functions.wait_and_click(self.driver, (AppiumBy.ACCESSIBILITY_ID, "buttonOK"), 20)

        Logger.logger_step(5, 'Closing the app')

        Functions.wait_and_click(self.driver, (AppiumBy.ACCESSIBILITY_ID, "buttonClose"), 20)

        Functions.wait_and_click(self.driver, (AppiumBy.ACCESSIBILITY_ID, "Close"), 20)


class Functions:
    @staticmethod
    def open_project(driver):
        Logger.logger_bold("Opening project")
        Functions.wait_and_click(driver, (AppiumBy.ACCESSIBILITY_ID, "buttonOK"), 25)

    @staticmethod
    def remove_test_project(driver):
        try:
            Logger.logger_bold("Closing the app and removing the test Project")
            driver.find_element(AppiumBy.NAME, "Close").click()
        except Exception as e:
            Logger.logger_warn(f"An error occurred while closing the app: {str(e)}")

    @staticmethod
    def wait_and_click(driver, locator, timeout):
        element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
        element.click()

    @staticmethod
    def wait_and_send_keys(driver, locator, timeout, text):
        element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator))
        element.send_keys(text)

    @staticmethod
    def wait_and_click_by_index(driver, locator, index, timeout):
        elements = WebDriverWait(driver, timeout).until(EC.presence_of_all_elements_located(locator), 0.1)
        if index < len(elements):
            elements[index].click()
        else:
            raise IndexError(f"Index {index} out of range. There are only {len(elements)} elements.")

    @staticmethod
    def validate_text_in_element(driver, expected_text, locator):
        element = driver.find_element(*locator)
        actual_text = element.text
        assert actual_text == expected_text, f"Expected '{expected_text}', but got '{actual_text}'"

    @staticmethod
    def double_click_element(element):
        x, y = element.location['x'], element.location['y']
        pyautogui.moveTo(x, y)
        pyautogui.doubleClick()

    @staticmethod
    def context_click_element(driver, locator):
        element = driver.find_element(*locator)
        x, y = element.location['x'], element.location['y']
        pyautogui.moveTo(x, y)
        pyautogui.rightClick()

    @staticmethod
    def drag_and_drop_element(source_element, target_element):
        source_x, source_y = source_element.location['x'], source_element.location['y']
        target_x, target_y = target_element.location['x'] + target_element.size['width'] // 2, target_element.location[
            'y'] + target_element.size['height'] // 2
        pyautogui.moveTo(source_x, source_y)
        pyautogui.mouseDown()
        time.sleep(0.5)
        pyautogui.moveTo(target_x, target_y)
        pyautogui.mouseUp()

    @staticmethod
    def resize_window(width, height):
        pyautogui.hotkey('alt', 'space')
        pyautogui.press('r')
        pyautogui.typewrite(['right', 'right', 'enter'])
        pyautogui.typewrite([str(width), 'tab', str(height), 'enter'])

    @staticmethod
    def uninstall_program(partial_program_name):
        try:
            subprocess.check_call(f'wmic product where "name like \'%{partial_program_name}%\'" call uninstall', shell=True)
            print(f"Program(s) starting with {partial_program_name} has been uninstalled.")
        except subprocess.CalledProcessError:
            print(f"No program starting with {partial_program_name} is installed or could not be uninstalled.")

    @staticmethod
    def install_program():
        try:
            Logger.logger_warn('Installing the latest verion of the app')
            cmd = [APP_EXE_PATH, "/S", "/v/qn"]
            result = subprocess.run(cmd, shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if result.returncode != 0:
                print(f"An error occurred during installation: {result.stderr.decode('utf-8')}")
            else:
                print("Installation was successful.")
        except Exception as e: 
            print(f"An exception occurred while installing the program: {e}")


class Logger:

    @staticmethod
    def logger_step(step_num, text):
        print(f"{PrintColors.PURPLE}{step_num}{' '}{text}{PrintColors.ENDCOLOR}")

    @staticmethod
    def logger_ok(text):
        print(f"{PrintColors.OKGREEN}{text}{PrintColors.ENDCOLOR}")

    @staticmethod
    def logger_warn(text):
        print(f"{PrintColors.WARNING}{text}{PrintColors.ENDCOLOR}")

    @staticmethod
    def logger_fail(text):
        print(f"{PrintColors.FAIL}{text}{PrintColors.ENDCOLOR}")

    @staticmethod
    def logger_bold(text):
        print(f"{PrintColors.BOLD}{text}{PrintColors.ENDCOLOR}")

    @staticmethod
    def logger_underline(text):
        print(f"{PrintColors.UNDERLINE}{text}{PrintColors.ENDCOLOR}")


class PrintColors:
    PURPLE = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDCOLOR = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


if __name__ == '__main__':
    unittest.main()