import pywinauto
from robot.api.deco import keyword
from pywinauto import application as pwa
from pywinauto.application import Application
from pywinauto import Desktop
import os
import subprocess


class SharepointSignIn():
    def __init__(self):
        self.app = None
        self.dlg = None
    def sharepointAuth(self,username,password):
        windows = Desktop(backend="uia").windows()
        avail_wind = [w.window_text() for w in windows]
        print(avail_wind)
        wind_bool = False
        for wind in avail_wind:
            if "windows security" in wind.lower():
                wind_bool = True
                break
        if wind_bool:
            main_app = Application(backend="uia").connect(title_re="Windows Security", control_type="Window")
            main_app_win = main_app.window(title_re="Windows Security", control_type="Window")
            # main_win_wrapper = main_app_win.set_focus()
            # main_win = main_app_win.child_window(title_re="Windows Security", control_type="Window")
            sign_in = main_app_win.child_window(title_re="", control_type="Pane")
            user_name = sign_in.child_window(class_name="TextBox", control_type="Edit")
            user_name.iface_value.SetValue(username)
            pass_word = sign_in.child_window(class_name="PasswordBox", control_type="Edit")
            pass_word.iface_value.SetValue(password)
            sign_in_button = main_app_win.child_window(title="OK", control_type="Button", visible_only=False)
            sign_in_button.invoke()



