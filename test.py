import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QComboBox, QPushButton, QLabel, QVBoxLayout, QWidget
from openpyxl import Workbook
from PyQt5.QtWidgets import QFileDialog, QMessageBox
import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QComboBox, QPushButton, QLabel, QVBoxLayout, QWidget, QFileDialog, QMessageBox
from openpyxl import Workbook

class VenuePlanner(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Venue Planner System')
        self.setGeometry(100, 100, 400, 400)

        # Initialize UI elements
        self.initUI()

        # Initialize data containers
        self.cambridge_data = None
        self.pearson_data = None
        self.subject_list = None
        self.time_data = None
        self.school_info = None
        self.exam_start_time = None
        self.add_materials = None

    def initUI(self):
        # Create layout
        layout = QVBoxLayout()

        # Zone Dropdown (disabled until all data is loaded)
        self.zone_label = QLabel("Select Zone:", self)
        self.zone_dropdown = QComboBox(self)
        self.zone_dropdown.addItems(["Zone 1", "Zone 2", "Zone 3", "Zone 4"])
        self.zone_dropdown.setEnabled(False)

        # Buttons for loading data
        self.load_cambridge_button = QPushButton('Load Cambridge Int. Data', self)
        self.load_cambridge_button.clicked.connect(self.load_cambridge_data)

        self.load_pearson_button = QPushButton('Load Pearson Edx. Data', self)
        self.load_pearson_button.clicked.connect(self.load_pearson_data)
        self.load_pearson_button.setEnabled(False)

        self.load_subject_list_button = QPushButton('Load Subject List', self)
        self.load_subject_list_button.clicked.connect(self.load_subject_list)
        self.load_subject_list_button.setEnabled(False)

        self.load_time_data_button = QPushButton('Load Time Data', self)
        self.load_time_data_button.clicked.connect(self.load_time_data)
        self.load_time_data_button.setEnabled(False)

        self.load_school_info_button = QPushButton('Load School Info', self)
        self.load_school_info_button.clicked.connect(self.load_school_info)
        self.load_school_info_button.setEnabled(False)

        self.load_exam_start_time_button = QPushButton('Load Exam Start Time', self)
        self.load_exam_start_time_button.clicked.connect(self.load_exam_start_time)
        self.load_exam_start_time_button.setEnabled(False)

        self.load_add_materials_button = QPushButton('Load Additional Materials', self)
        self.load_add_materials_button.clicked.connect(self.load_add_materials)
        self.load_add_materials_button.setEnabled(False)

        # Submit Button (disabled until all data is loaded)
        self.submit_button = QPushButton('Submit', self)
        self.submit_button.clicked.connect(self.create_venue_plan)
        self.submit_button.setEnabled(False)

        # Add widgets to the layout
        layout.addWidget(self.zone_label)
        layout.addWidget(self.zone_dropdown)
        layout.addWidget(self.load_cambridge_button)
        layout.addWidget(self.load_pearson_button)
        layout.addWidget(self.load_subject_list_button)
        layout.addWidget(self.load_time_data_button)
        layout.addWidget(self.load_school_info_button)
        layout.addWidget(self.load_exam_start_time_button)
        layout.addWidget(self.load_add_materials_button)
        layout.addWidget(self.submit_button)

        # Set layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def load_cambridge_data(self):
        cambridge_file, _ = QFileDialog.getOpenFileName(self, "Select Cambridge Int. Excel File", "", "Excel Files (*.xlsx *.xls)")
        if cambridge_file:
            self.cambridge_data = pd.read_excel(cambridge_file)
            QMessageBox.information(self, "Loaded", "Cambridge Int. data loaded successfully.")
            self.load_pearson_button.setEnabled(True)

    def load_pearson_data(self):
        pearson_file, _ = QFileDialog.getOpenFileName(self, "Select Pearson Edx. Excel File", "", "Excel Files (*.xlsx *.xls)")
        if pearson_file:
            self.pearson_data = pd.read_excel(pearson_file)
            QMessageBox.information(self, "Loaded", "Pearson Edx. data loaded successfully.")
            self.load_subject_list_button.setEnabled(True)

    def load_subject_list(self):
        subject_list_file, _ = QFileDialog.getOpenFileName(self, "Select Subject List Excel File", "", "Excel Files (*.xlsx *.xls)")
        if subject_list_file:
            self.subject_list = pd.read_excel(subject_list_file)
            QMessageBox.information(self, "Loaded", "Subject List loaded successfully.")
            self.load_time_data_button.setEnabled(True)

    def load_time_data(self):
        time_data_file, _ = QFileDialog.getOpenFileName(self, "Select Time Data Excel File", "", "Excel Files (*.xlsx *.xls)")
        if time_data_file:
            self.time_data = pd.read_excel(time_data_file)
            QMessageBox.information(self, "Loaded", "Time Data loaded successfully.")
            self.load_school_info_button.setEnabled(True)

    def load_school_info(self):
        school_info_file, _ = QFileDialog.getOpenFileName(self, "Select School Info Excel File", "", "Excel Files (*.xlsx *.xls)")
        if school_info_file:
            self.school_info = pd.read_excel(school_info_file)
            QMessageBox.information(self, "Loaded", "School Info loaded successfully.")
            self.load_exam_start_time_button.setEnabled(True)

    def load_exam_start_time(self):
        exam_start_time_file, _ = QFileDialog.getOpenFileName(self, "Select Exam Start Time Excel File", "", "Excel Files (*.xlsx *.xls)")
        if exam_start_time_file:
            self.exam_start_time = pd.read_excel(exam_start_time_file)
            QMessageBox.information(self, "Loaded", "Exam Start Time loaded successfully.")
            self.load_add_materials_button.setEnabled(True)

    def load_add_materials(self):
        add_materials_file, _ = QFileDialog.getOpenFileName(self, "Select Additional Materials Excel File", "", "Excel Files (*.xlsx *.xls)")
        if add_materials_file:
            self.add_materials = pd.read_excel(add_materials_file)
            QMessageBox.information(self, "Loaded", "Additional Materials loaded successfully.")
            self.zone_dropdown.setEnabled(True)
            self.submit_button.setEnabled(True)

    def create_venue_plan(self):
        # Generate a venue plan based on the loaded data and selected zone
        selected_zone = self.zone_dropdown.currentText()
        print(f"Creating venue plan for {selected_zone}")

        # Example merging data to create a venue plan
        merged_data = pd.merge(self.subject_list, self.school_info, how="left", on="School")
        merged_data['Start Time'] = self.exam_start_time['Start Time']

        # Output file creation using openpyxl
        wb = Workbook()
        ws = wb.active
        ws.title = "Venue Plan System"

        # Write headers
        headers = ["School", "Subject", "Zone", "Start Time", "Additional Materials"]
        ws.append(headers)

        # Add filtered data into the workbook
        for index, row in merged_data.iterrows():
            ws.append([row['School'], row['Subject'], selected_zone, row['Start Time'], "Materials Needed"])

        # Save workbook
        output_file = f"Venue_Plan_{selected_zone}.xlsx"
        wb.save(output_file)

        QMessageBox.information(self, "Success", f"Venue plan saved as {output_file}")
        print(f"Venue plan created and saved as {output_file}")

# Main application execution
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = VenuePlanner()
    window.show()
    sys.exit(app.exec_())
