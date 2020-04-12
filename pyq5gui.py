import PyQt5
from PyQt5 import QtCore, QtGui, uic, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QWidget
from PyQt5 import QtWidgets, uic
import sys
import json
import win32com.client
xl = win32com.client.Dispatch("Excel.Application")

defalt_config_folder = {"conf_txt_link_motable":"","conf_txt_link_termcode":"","conf_txt_link_pepfile":"","conf_txt_link_hsdatabase":""}
if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)

if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    PyQt5.QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        self.setFixedSize(600, 360)
        uic.loadUi('./assets/GUI.ui', self)

        #Logical group ComboBox
        self.select_logical_group = self.findChild(QtWidgets.QComboBox,'select_logical_group')
        self.select_logical_group.currentTextChanged.connect(self.change_logical_group)
        #Refresh button
        self.btt_load_data = self.findChild(QtWidgets.QPushButton, 'btt_load_data') # Find the button
        self.btt_load_data.clicked.connect(self.load_data)
        #Load file button
        self.btt_load_file = self.findChild(QtWidgets.QPushButton, 'btt_load_file') # Find the button
        self.btt_load_file.clicked.connect(self.load_file)
        #Start button
        self.btt_start_run = self.findChild(QtWidgets.QPushButton, 'btt_start_run') # Find the button
        self.btt_start_run.clicked.connect(self.start_run)
        #Rule Bar
        self.bar_rule_abc  = self.findChild(QtWidgets.QProgressBar, 'bar_rule_abc')
        self.bar_rule_d  = self.findChild(QtWidgets.QProgressBar, 'bar_rule_d')
        self.bar_no_rule  = self.findChild(QtWidgets.QProgressBar, 'bar_no_rule')
        self.bar_classification  = self.findChild(QtWidgets.QProgressBar, 'bar_classification')
        #Text box
        self.txt_link_to_file = self.findChild(QtWidgets.QLineEdit,'txt_link_to_file')
        #Text Lable
        self.txt_status = self.findChild(QtWidgets.QLabel,'txt_status')

#Configuration Tap############################
        #Reset to default button
        self.conf_btt_set_todefault = self.findChild(QtWidgets.QPushButton, 'conf_btt_set_todefault') # Find the button
        self.conf_btt_set_todefault.clicked.connect(self.conf_set_todefault)

        #MO table button to pickup forler:
        self.conf_btt_set_motable = self.findChild(QtWidgets.QPushButton, 'conf_btt_set_motable') # Find the button
        self.conf_btt_set_motable.clicked.connect(self.conf_set_motable)

        #Term Code button to pickup forler:
        self.conf_btt_set_termcode = self.findChild(QtWidgets.QPushButton, 'conf_btt_set_termcode') # Find the button
        self.conf_btt_set_termcode.clicked.connect(self.conf_set_termcode)

        #PEP file button to pickup forler:
        self.conf_btt_set_pepfile = self.findChild(QtWidgets.QPushButton, 'conf_btt_set_pepfile') # Find the button
        self.conf_btt_set_pepfile.clicked.connect(self.conf_set_pepfile)

        #HScode data button to pickup forler:
        self.conf_btt_set_hsdatabase = self.findChild(QtWidgets.QPushButton, 'conf_btt_set_hsdatabase') # Find the button
        self.conf_btt_set_hsdatabase.clicked.connect(self.conf_set_hsdatabase)

        #Text box
        self.conf_txt_link_motable = self.findChild(QtWidgets.QLineEdit,'conf_txt_link_motable')
        self.conf_txt_link_termcode = self.findChild(QtWidgets.QLineEdit,'conf_txt_link_termcode')
        self.conf_txt_link_pepfile = self.findChild(QtWidgets.QLineEdit,'conf_txt_link_pepfile')
        self.conf_txt_link_hsdatabase = self.findChild(QtWidgets.QLineEdit,'conf_txt_link_hsdatabase')





#Run before gui active#########################################
        # Load config folder :
        try:
            f = open("config_folder.txt", "r")
            folder_config = json.loads(f.readline())
            f.close()
        except:
            f = open("config_folder.txt", "w")
            f.write(json.dumps(defalt_config_folder))
            f.close()
            folder_config = defalt_config_folder
        self.conf_txt_link_motable.setText(folder_config["conf_txt_link_motable"])
        self.conf_txt_link_termcode.setText(folder_config["conf_txt_link_termcode"])
        self.conf_txt_link_pepfile.setText(folder_config["conf_txt_link_pepfile"])
        self.conf_txt_link_hsdatabase.setText(folder_config["conf_txt_link_hsdatabase"])
        # End load config folder
        # Get Ative workbook name
        try:
            self.txt_link_to_file.setText(xl.ActiveWorkbook.fullname)
        except:
            self.txt_link_to_file.setText('no file found')

        # End get active workbook name
#end Run before gui active#####################################
        self.show()

    def load_data(self):
        # This is executed when the button is pressed
        #self.pro_bar.setValue(50)
        print('you clickk refresh button')
        return
    def change_logical_group(self,value):
        print('you change logical group: ' +str(value))
        return
        btt_load_file
    def load_file(self):
        #file = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        file = str(QFileDialog.getOpenFileName(self, 'Select file')[0])
        self.txt_link_to_file.setText(file)
        print('you link: ' + str(file) )
        return
    def start_run(self):
        print('you clickk btt_start_run button')
        status = 'Running: {} /{} - {} part/s - estimate finished time: {} mins'
        part_done = str(100)
        total_part = str(500)
        run_speed = str(20)
        time_left = str(30)
        percent_done = 30
        self.bar_classification.setValue(percent_done)
        self.txt_status.setText(status.format(part_done,total_part,run_speed,time_left))
        return
    def conf_set_todefault(self):
        #print('you clickk conf_reset_todefault button')
        f = open("config_folder.txt", "r")
        folder_config = json.loads(f.readline())

        folder_config["conf_txt_link_motable"] = self.conf_txt_link_motable.text()
        folder_config["conf_txt_link_termcode"] = self.conf_txt_link_termcode.text()
        folder_config["conf_txt_link_pepfile"] = self.conf_txt_link_pepfile.text()
        folder_config["conf_txt_link_hsdatabase"] = self.conf_txt_link_hsdatabase.text()
        f = open("config_folder.txt", "w")
        f.write(json.dumps(folder_config))
        f.close()
        return
#get file or table link to save config file###################
    def conf_set_motable(self):
        folder_link = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.conf_txt_link_motable.setText(folder_link)
        return
    def conf_set_termcode(self):
        folder_link = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.conf_txt_link_termcode.setText(folder_link)
        return
    def conf_set_pepfile(self):
        folder_link = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.conf_txt_link_pepfile.setText(folder_link)
        return
    def conf_set_hsdatabase(self):
        folder_link = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.conf_txt_link_hsdatabase.setText(folder_link)
        return
#END get file or table link to save config file###################
    def get_folder_link(seft):

        return str(QFileDialog.getExistingDirectory(self, "Select Directory"))

    def get_file_link(seft):

        return str(QFileDialog.getOpenFileName(self, 'Select file')[0])

app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()
