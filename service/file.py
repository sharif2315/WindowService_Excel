import socket
import win32serviceutil
import servicemanager
import win32event
import win32service
import os
import pyodbc
import shutil
import time
import random
from pathlib import Path
#from SMWinservice import SMWinservice


class SMWinservice(win32serviceutil.ServiceFramework):
    '''Base class to create winservice in Python'''

    _svc_name_ = 'pythonService'
    _svc_display_name_ = 'Python Service'
    _svc_description_ = 'Python Service Description'

    @classmethod
    def parse_command_line(cls):
        '''
        ClassMethod to parse the command line
        '''
        win32serviceutil.HandleCommandLine(cls)

    def __init__(self, args):
        '''
        Constructor of the winservice
        '''
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        '''
        Called when the service is asked to stop
        '''
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        '''
        Called when the service is asked to start
        '''
        self.start()
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def start(self):
        '''
        Override to add logic before the start
        eg. running condition
        '''
        pass

    def stop(self):
        '''
        Override to add logic before the stop
        eg. invalidating running condition
        '''
        pass

    def main(self):
        '''
        Main class to be ovverridden to add logic
        '''
        pass


class PythonCornerExample(SMWinservice):
    _svc_name_ = "ExcelFileMoverService"
    _svc_display_name_ = "Excel File Mover Service"
    _svc_description_ = "Moves xlsx files into archive folder"

    def movefiles(self):
        mypath = r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test'
        myarchive = r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test\archive'
        filelist = []
        # reading directory
        for dirpath, dirnames, filenames in os.walk(mypath):
            for filename in filenames:
                # Reading directory for only xlsx files which are in dir test
                if (filename.lower().endswith('.xlsx')) and (dirpath == mypath):
                    filepath = dirpath + '\\' + filename
                    filelist.append(filepath)
                else:
                    continue

        for file in filelist:
            filename = os.path.basename(file)
            shutil.move( file, str(myarchive + '\\' + filename) )
            #print('File: ' + str(file) + ' has been successfully moved to archive folder')
            print('File: ' + str(filename) + ' has been successfully moved to archive folder')

    def start(self):
        self.isrunning = True

    def stop(self):
        self.isrunning = False

    def main(self):
        #
        #movefiles()
        #i = 0
        while self.isrunning:
            try:
                mypath = r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test'
                myarchive = r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test\archive'
                filelist = []
                # reading directory
                for dirpath, dirnames, filenames in os.walk(mypath):
                    for filename in filenames:
                        # Reading directory for only xlsx files which are in dir test
                        if (filename.lower().endswith('.xlsx')) and (dirpath == mypath):
                            filepath = dirpath + '\\' + filename
                            filelist.append(filepath)
                        else:
                            continue

                for file in filelist:
                    filename = os.path.basename(file)
                    shutil.move( file, str(myarchive + '\\' + filename) )
                    #print('File: ' + str(file) + ' has been successfully moved to archive folder')
                    #print('File: ' + str(filename) + ' has been successfully moved to archive folder')


                # movefiles(self)
                # open(r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test\myfile.txt', 'x')
            # except NameError:
            #     open(r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test\log.txt', 'x')
            #     f = open(r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test\log.txt', 'a')
            #     f.write('created myfile.txt and then created log.txt')
            #     f.close()
            except Exception as e:

                open(r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test\log.txt', 'x')
                f = open(r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test\log.txt', 'a')
                f.write(str(e))
                f.close()

                time.sleep(5)
                stop(self)


if __name__ == '__main__':
    PythonCornerExample.parse_command_line()
