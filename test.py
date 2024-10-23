import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QMessageBox
from openpyxl import Workbook

class VenuePlanner(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Venue Planner')
        self.setGeometry(100, 100, 400, 400)

        # Initialize UI elements
        self.initUI()

        # Initialize data containers
        self.cambridge_data = None
        self.pearson_data = None
        self.subject_list = None
        self.school_info = None

    def initUI(self):
        # Create layout
        layout = QVBoxLayout()

        # Buttons for loading data
        self.load_cambridge_button = QPushButton('Load Cambridge Int. Data', self)
        self.load_cambridge_button.clicked.connect(self.load_cambridge_data)

        self.load_pearson_button = QPushButton('Load Pearson Edx. Data', self)
        self.load_pearson_button.clicked.connect(self.load_pearson_data)
        self.load_pearson_button.setEnabled(False)

        self.load_subject_list_button = QPushButton('Load Subject List', self)
        self.load_subject_list_button.clicked.connect(self.load_subject_list)
        self.load_subject_list_button.setEnabled(False)

        self.load_school_info_button = QPushButton('Load School Info', self)
        self.load_school_info_button.clicked.connect(self.load_school_info)
        self.load_school_info_button.setEnabled(False)

        # Submit Button (disabled until all data is loaded)
        self.submit_button = QPushButton('Submit', self)
        self.submit_button.clicked.connect(self.create_venue_plan)
        self.submit_button.setEnabled(False)

        # Add widgets to the layout
        layout.addWidget(self.load_cambridge_button)
        layout.addWidget(self.load_pearson_button)
        layout.addWidget(self.load_subject_list_button)
        layout.addWidget(self.load_school_info_button)
        layout.addWidget(self.submit_button)

        # Set layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def load_cambridge_data(self):
        cambridge_file, _ = QFileDialog.getOpenFileName(self, "Select Cambridge Int. Excel File", "", "Excel Files (*.xlsx *.xls)")
        if cambridge_file:
            self.cambridge_data = pd.read_excel(cambridge_file)
            print("Cambridge Data Columns:", self.cambridge_data.columns)  # Debugging: Check column names
            QMessageBox.information(self, "Loaded", "Cambridge Int. data loaded successfully.")
            self.load_pearson_button.setEnabled(True)

    def load_pearson_data(self):
        pearson_file, _ = QFileDialog.getOpenFileName(self, "Select Pearson Edx. Excel File", "", "Excel Files (*.xlsx *.xls)")
        if pearson_file:
            self.pearson_data = pd.read_excel(pearson_file)
            print("Pearson Data Columns:", self.pearson_data.columns)  # Debugging: Check column names
            QMessageBox.information(self, "Loaded", "Pearson Edx. data loaded successfully.")
            self.load_subject_list_button.setEnabled(True)

    def load_subject_list(self):
        subject_list_file, _ = QFileDialog.getOpenFileName(self, "Select Subject List Excel File", "", "Excel Files (*.xlsx *.xls)")
        if subject_list_file:
            self.subject_list = pd.read_excel(subject_list_file)
            print("Subject List Columns:", self.subject_list.columns)  # Debugging: Check column names
            QMessageBox.information(self, "Loaded", "Subject List loaded successfully.")
            self.load_school_info_button.setEnabled(True)

    def load_school_info(self):
        school_info_file, _ = QFileDialog.getOpenFileName(self, "Select School Info Excel File", "", "Excel Files (*.xlsx *.xls)")
        if school_info_file:
            self.school_info = pd.read_excel(school_info_file)
            print("School Info Columns:", self.school_info.columns)  # Debugging: Check column names
            QMessageBox.information(self, "Loaded", "School Info loaded successfully.")
            self.submit_button.setEnabled(True)

    def create_venue_plan(self):
        # Ensure necessary columns exist as expected
        required_cambridge_cols = ['Subject/Component Code', 'Centre Number', 'Candidate Number', 'Date of Birth', 'Gender', 'Mobile Phone', 'Email Address']
        required_pearson_cols = ['Subject/Component Code', 'Centre Number', 'Candidate Number', 'Date of Birth', 'Gender', 'Mobile Phone', 'Email Address']
        required_subject_cols = ['Subject/Component Code', 'Type']
        required_school_cols = ['Centre Number', 'School Name', 'Centre type', 'Zone', 'Location']

        # Convert 'Centre Number', 'Mobile Phone', and 'Email Address' to string in both DataFrames
        self.cambridge_data['Centre Number'] = self.cambridge_data['Centre Number'].astype(str)
        self.cambridge_data['Mobile Phone'] = self.cambridge_data['Mobile Phone'].astype(str)
        self.cambridge_data['Email Address'] = self.cambridge_data['Email Address'].astype(str)

        self.pearson_data['Centre Number'] = self.pearson_data['Centre Number'].astype(str)
        self.pearson_data['Mobile Phone'] = self.pearson_data['Mobile Phone'].astype(str)
        self.pearson_data['Email Address'] = self.pearson_data['Email Address'].astype(str)

        # Convert date columns to datetime format
        self.cambridge_data['Date of Birth'] = pd.to_datetime(self.cambridge_data['Date of Birth'], errors='coerce')
        self.pearson_data['Date of Birth'] = pd.to_datetime(self.pearson_data['Date of Birth'], errors='coerce')

        # Merging Cambridge and Pearson data on the common columns
        merged_data = pd.merge(
            self.cambridge_data[required_cambridge_cols],
            self.pearson_data[required_pearson_cols],
            on=["Subject/Component Code", "Centre Number", "Candidate Number", "Date of Birth", "Gender", "Mobile Phone", "Email Address"],
            how="outer"
        )

        # Merging with subject list to add 'Type'
        merged_data = pd.merge(merged_data, self.subject_list[required_subject_cols], on="Subject/Component Code", how="left")

        # Check if required columns are present in school info before merging
        if not all(col in self.school_info.columns for col in required_school_cols):
            missing_cols = [col for col in required_school_cols if col not in self.school_info.columns]
            QMessageBox.warning(self, "Error", f"Missing columns in School Info data: {', '.join(missing_cols)}")
            return

        # Merging with school info to add 'School Name', 'Centre Type', 'Zone', 'Location'
        merged_data = pd.merge(merged_data, self.school_info[required_school_cols], on="Centre Number", how="left")

        # Sort data by Centre Number and Subject/Component Code
        sorted_data = merged_data.sort_values(by=["Centre Number", "Subject/Component Code"])

        # Create Excel output using openpyxl
        wb = Workbook()
        ws = wb.active
        ws.title = "Venue Planner System"

        # Write headers
        ws.append(sorted_data.columns.tolist())

        # Write data
        for _, row in sorted_data.iterrows():
            ws.append(row.tolist())

        # Save workbook
        output_file = "Venue_Plan.xlsx"
        wb.save(output_file)

        QMessageBox.information(self, "Success", f"Venue plan saved as {output_file}")
        print(f"Venue plan created and saved as {output_file}")


# Main application execution
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = VenuePlanner()
    window.show()
    sys.exit(app.exec_())

