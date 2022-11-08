import pywinauto
from robot.api.deco import keyword
from pywinauto import application as pwa
from pywinauto.application import Application
from pywinauto import Desktop
import os
import subprocess
import time

#def exec_sp(qualityPath):
    #fileUploadWindow().sharepointfileupload(qualityPath)
class fileUploadWindow:
    #def __init__(self):
        #self.app = None
        #self.dlg = None

    #@keyword("Windows Authentication")
    #def sharepointfileupload(self,username, password, title_name):
    def getPath(self,ini_dir):#\\acdev01\3M_CAC\ERP_Quality_Review\sharepointUpload
        paths=list()
        fileleng=list()
        for direct in os.listdir(ini_dir):
            paths.append(os.path.join(ini_dir,direct))
            fileleng.append(len(os.listdir(os.path.join(ini_dir,direct))))
        return {"fpaths":''.join(paths),"fileLength":str(fileleng[0])}
    def sharepointfileupload(self,qualityPath):
        windows = Desktop(backend="uia").windows()
        print([w.window_text() for w in windows])
        avail_wind = [w.window_text() for w in windows]
        MainApplicationName="Project Alpine Execution"       
        firstexe=self.getPath(qualityPath)#r"\\acdev01\3M_CAC\ERP_Quality_Review\sharepointUpload")#"\\\\acdev01\\3M_CAC\\AOMS_user_role"#"\\acdev01\3M_CAC\AOMS_user_role" 
        path=firstexe.get("fpaths")
        flength=firstexe.get("fileLength")
        main_app = Application(backend="uia").connect(title_re=".*%s.*" % MainApplicationName, control_type="Window")
        app_dialog =main_app.top_window()
        app_dialog.set_focus()
        app_child = app_dialog.child_window(title="Select folder to upload", control_type="Window")
        #app_child.print_control_identifiers()
        app_child_1 = app_child.child_window(title="Folder:", auto_id="1152", control_type="Edit")
        app_child_1.iface_value.SetValue(path)       
        UploadButton=app_child.child_window(title="Upload", auto_id="1", control_type="Button")
        UploadButton.invoke()
        #app_dialog.set_focus()
        time.sleep(5)
        permission=app_dialog.child_window(title="Upload {filecount} files to this site?".format(filecount=str(flength)),control_type="Window")
        #permission.print_control_identifiers()
        Uploadgo=permission.child_window(title="Upload", control_type="Button")
        Uploadgo.invoke()
#exec_sp()
