import os
import time
import unittest
import warnings
import winshell
from appium import webdriver
from appium.webdriver.common.appiumby import AppiumBy
from appium.options.windows import WindowsOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import socket
import ctypes
from appium.webdriver.appium_service import AppiumService

"""
Install dependencies:
node 18 or 20
npm install -g appium@2.2.3
appium driver install --source=npm appium-windows-driver
pip install Appium-Python-Client
pip install selenium
pip install winshell
"""

# Constants
SHORTCUT_PATH = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Program\\Application.lnk"
EXE_DIRECTORY = r'C:\Program Files\Application'
DEFAULT_PORT = 4725

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

    @classmethod
    def setUpClass(cls):
        warnings.filterwarnings("ignore")

        Logger.logger_warn('Killing all instances of node.exe and WinAppDriver.exe to avoid interference of any running Appium sessions')

        os.system('taskkill /f /im node.exe')
        os.system('taskkill /f /im WinAppDriver.exe')

        time.sleep(2)

        Logger.logger_warn("All instances of 'node' and 'WinAppDriver' killed, starting Appium server, please wait")

        cls.appium_server.start()

        # Parsing shortcut so that if the app needs to be launched with some arguments, they will be parsed automatically and the app will be launched using them
        shortcut = winshell.shortcut(SHORTCUT_PATH)
        target_path = shortcut.path
        arguments = shortcut.arguments
        Logger.logger_underline(
            f"Launching the app from URI: {target_path} with the following arguments (if any): {arguments}")

        warnings.filterwarnings("ignore")

        ctypes.windll.shell32.ShellExecuteW(None, "runas", target_path, arguments, EXE_DIRECTORY, 1)

    def setUp(self) -> None:
        appium_server_url = f"http://127.0.0.1:{self.appium_server.port}/wd/hub"
        options = WindowsOptions()
        options.experimental_webdriver = True
        options.app = "Root"
        options.platform_name = "Windows"
        self.driver = webdriver.Remote(appium_server_url, options=options)
        self.wait = WebDriverWait(self.driver, 10)

    def tearDown(self) -> None:
        if self.driver is not None:
            try:
                Logger.logger_bold("Closing the app after test")
                self.driver.find_element(AppiumBy.NAME, "Close").click()
            except Exception as e:
                Logger.logger_warn(f"An error occurred while closing the app: {str(e)}")

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

    def test_do_something(self):
        Functions.create_new_project(self.driver, "TestProject", "123456")

        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "New Area"), 10)

        Logger.logger_step(2, "Dropping new area at location")

        new_entry_xpath = "//*[@AutomationId=\"MainForm\"]/Pane[@AutomationId=\"m_newEntry\"]"
        Functions.wait_and_click(self.driver, (AppiumBy.XPATH, new_entry_xpath), 10)

        Logger.logger_step(3, "Naming the new entry")
        Functions.wait_and_send_keys(self.driver, (AppiumBy.NAME, "New Area"), 10, 'Hello Kitty :)')
        time.sleep(0.2)
        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "OK"), 10)

        Logger.logger_step(4, "Validating the entry name")

        Functions.validate_text_in_element(self.driver, "Hello Kitty :)", (AppiumBy.NAME, "Name Row 0"))

        Logger.logger_step(5, "Dial some number, cancel dialing process")

        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "System"), 10)

        dropdown_menu_locator = (AppiumBy.NAME, "SystemDropDown")

        Functions.wait_and_click_by_index(self.driver, dropdown_menu_locator, 0, 10)

        Functions.wait_and_send_keys(self.driver, (AppiumBy.NAME, "Dialup"), 10, "05356789574")

        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "Dial"), 10)

        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "Cancel"), 10)

        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "Hangup"), 10)

        Functions.wait_and_click(self.driver, (AppiumBy.NAME, "Close"), 10)
        
        pass

class Functions:
    @staticmethod
    def create_new_project(driver, project_name, password):
        Logger.logger_step(1, "Creating a new project")
        Functions.wait_and_click(driver, (AppiumBy.NAME, "New project..."), 10)
        driver.find_element(AppiumBy.NAME, "Name:").send_keys(project_name)
        driver.find_element(AppiumBy.NAME, "Password:").send_keys(password)
        driver.find_element(AppiumBy.NAME, "Confirm password:").send_keys(password)
        Functions.wait_and_click(driver, (AppiumBy.NAME, "OK"), 10)

    @staticmethod
    def wait_and_click(driver, locator, timeout):
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator)).click()

    @staticmethod
    def wait_clickable_and_click(driver, locator, timeout):
        WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(locator)).click()
        

    @staticmethod
    def wait_and_send_keys(driver, locator, timeout, text):
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located(locator)).send_keys(text)

    @staticmethod
    def wait_and_click_by_index(driver, locator, index, timeout):
        elements = WebDriverWait(driver, timeout).until(EC.presence_of_all_elements_located(locator))
        if index < len(elements):
            elements[index].click()
        else:
            raise IndexError(f"Index {index} out of range. There are only {len(elements)} elements.")

    @staticmethod
    def validate_text_in_element(driver, expected_text, locator):
        element = driver.find_element(*locator)
        actual_text = element.text
        assert actual_text == expected_text, f"Expected '{expected_text}', but got '{actual_text}'"


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
