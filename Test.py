import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QFrame, QTextEdit, QFileDialog, QMessageBox,
                             QScrollArea, QSplitter)
from PyQt5.QtGui import QFont, QIcon, QFontDatabase
from PyQt5.QtCore import Qt, QSize
from Datainput import DataInputTab
from math import pi, sqrt
from casing import DbCalculator

class Colors:
    PRIMARY = "#2b2b2b"
    SECONDARY = "#3b3b3b"
    ACCENT = "#4CAF50"
    TEXT = "#ffffff"
    SUBTLE = "#555555"
    HIGHLIGHT = "#45a049"

class WellDataApp(QWidget):
    STYLE_SHEET = """
    QWidget {
        background-color: %(PRIMARY)s;
        color: %(TEXT)s;
        font-family: 'Roboto', sans-serif;
    }
    QPushButton {
        background-color: %(ACCENT)s;
        color: %(TEXT)s;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        font-size: 16px;
        font-weight: bold;
    }
    QPushButton:hover {
        background-color: %(HIGHLIGHT)s;
    }
    QPushButton:pressed {
        background-color: #3d8b40;
    }
    QLabel {
        color: %(TEXT)s;
    }
    QTextEdit {
        background-color: %(SECONDARY)s;
        color: %(TEXT)s;
        border: 1px solid %(SUBTLE)s;
        border-radius: 5px;
        padding: 8px;
    }
    QScrollArea, QSplitter::handle {
        background-color: %(PRIMARY)s;
    }
    QSplitter::handle:horizontal {
        width: 4px;
    }
    QSplitter::handle:vertical {
        height: 4px;
    }
    QSplitter::handle:pressed {
        background-color: %(ACCENT)s;
    }
    QTabWidget::pane {
        border: 1px solid %(SUBTLE)s;
        background: %(SECONDARY)s;
    }
    QTabWidget::tab-bar {
        left: 5px;
    }
    QTabBar::tab {
        background: %(PRIMARY)s;
        border: 1px solid %(SUBTLE)s;
        padding: 5px;
        color: %(TEXT)s;
    }
    QTabBar::tab:selected {
        background: %(ACCENT)s;
    }
    QTabBar::tab:hover {
        background: %(HIGHLIGHT)s;
    }
    """ % Colors.__dict__

    def __init__(self):
        super().__init__()
        self.drill_collar_diameters_mm = []
        self.data_input_tab = DataInputTab()
        self.casing_tab = DbCalculator()
        self.df = None
        self.additional_columns = []
        self.nearest_bit_sizes = []
        self.drill_pipe_data = {}
        self.setup_ui()
        
    def setup_ui(self):
        QFontDatabase.addApplicationFont("fonts/Roboto-Bold.ttf")
        QFontDatabase.addApplicationFont("fonts/Roboto-Regular.ttf")
        self.setFont(QFont("Roboto", 10))
        self.setStyleSheet(self.STYLE_SHEET)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        self.setup_top_bar(main_layout)
        self.setup_main_content(main_layout)
        
    def setup_top_bar(self, main_layout):
        top_bar = QHBoxLayout()
        
        self.upload_file_btn = self.create_button("Upload Drill Collar Table", "icons/upload.png", self.upload_excel_file)
        self.calculate_drill_collar_btn = self.create_button("Calculate Drill Collar", "icons/drill.png", self.calculate_drill_collar)
        self.calculate_data_btn = self.create_button("Calculate Data", "icons/calculate.png", self.calculate_and_display)
        
        top_bar.addWidget(self.upload_file_btn)
        top_bar.addWidget(self.calculate_drill_collar_btn)
        top_bar.addWidget(self.calculate_data_btn)
        top_bar.addStretch()
        
        main_layout.addLayout(top_bar)
        
    def setup_main_content(self, main_layout):
        splitter = QSplitter(Qt.Horizontal)
        
        left_panel = self.create_panel("Calculations")
        right_panel = self.create_panel("Results")
        
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([int(self.width() * 0.4), int(self.width() * 0.6)])
        
        self.calculation_text = QTextEdit()
        self.calculation_text.setReadOnly(True)
        left_panel.layout().addWidget(self.calculation_text)
        
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        right_panel.layout().addWidget(self.result_text)
        
        main_layout.addWidget(splitter)

    def create_button(self, text, icon_path, connection):
        btn = QPushButton(text)
        btn.setIcon(QIcon(icon_path))
        btn.setIconSize(QSize(24, 24))
        btn.clicked.connect(connection)
        return btn

    def create_panel(self, title):
        panel = QFrame()
        panel.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        layout = QVBoxLayout(panel)
        
        label = QLabel(title)
        label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(label)
        
        return panel

    def upload_excel_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Drill Collar Table Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        
        if file_path:
            try:
                self.load_drill_collar_data(file_path)
                QMessageBox.information(self, "Success", "Drill Collar Table loaded successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load Excel file. Error: {e}")
        else:
            QMessageBox.warning(self, "No File Selected", "Please select an Excel file.")

    def load_drill_collar_data(self, file_path):
        try:
            self.df = pd.read_excel(file_path, sheet_name='sheet1')
            self.df.columns = self.df.columns.str.strip()
            self.df.rename(columns={
                'Drill pipe Metal grade': 'Drill pipe Metal grade',
                'Minimum tensile strength(psi)': 'Minimum tensile strength(psi)',
                'Minimum tensile strength(mpi)': 'Minimum tensile strength(mpi)'
            }, inplace=True)
            
            self.drill_collar_diameters_mm = pd.to_numeric(self.df['Drilling collars outer diameter'], errors='coerce').dropna().values
            self.drill_collar_diameters_mm = self.drill_collar_diameters_mm.astype(float)
            
            self.additional_columns = [
                'Outer diameter', 'AP', 'AIP', 'Mp', 'qp', 'b', 'γ'
            ]
            
            self.df['qp'] = self.df['qp'].astype(float)
            self.df['γ'] = self.df['γ'].astype(float)
            
            self.drill_pipe_data = {
                'Drill pipe Metal grade': self.df['Drill pipe Metal grade'].unique().tolist(),
                'Minimum tensile strength(psi)': self.df['Minimum tensile strength(psi)'].unique().tolist(),
                'Minimum tensile strength(mpi)': self.df['Minimum tensile strength(mpi)'].unique().tolist()
            }
            
        except Exception as e:
            raise Exception(f"Error loading drill collar data: {e}")

    def nearest_drill_collar(self, value):
        if len(self.drill_collar_diameters_mm) == 0:
            return None
        
        idx = (np.abs(self.drill_collar_diameters_mm - value)).argmin()
        return self.drill_collar_diameters_mm[idx]
    
    def calculate_drill_collar(self):
        if len(self.drill_collar_diameters_mm) == 0:
            self.result_text.setHtml("<p style='color: #F44747;'>No Drill Collar Table loaded. Please upload an Excel file.</p>")
            return
        
        initial_dcsg, at_head_values, nearest_bit_sizes, _ = self.casing_tab.get_all_dcsg_values()
        if initial_dcsg and at_head_values and nearest_bit_sizes:
            self.display_drill_collar_results(initial_dcsg, at_head_values, nearest_bit_sizes)
        else:
            self.result_text.setHtml("<p style='color: #F44747;'>Unable to retrieve Dcsg values. Please check the Casing tab.</p>")

    def display_drill_collar_results(self, initial_dcsg, at_head_values, nearest_bit_sizes):
        self.nearest_bit_sizes = nearest_bit_sizes
        html_result = """
        <style>
            table {width: 100%; border-collapse: collapse; margin-bottom: 20px;}
            th, td {border: 1px solid #555555; padding: 8px; text-align: left;}
            th {background-color: #4CAF50; color: #FFFFFF;}
            tr:nth-child(even) {background-color: #3b3b3b;}
        </style>
        <table>
            <tr>
                <th>Section</th>
                <th>At head (Dcsg)</th>
                <th>Nearest Bit Size</th>
                <th>Drill Collars</th>
            </tr>
        """
        
        all_values = list(zip([initial_dcsg] + at_head_values, nearest_bit_sizes))
        total_iterations = len(all_values)
        
        for i, (at_head, bit_size) in enumerate(all_values):
            drill_collar = 2 * float(at_head) - bit_size
            nearest_drill_collar = self.nearest_drill_collar(drill_collar)

            if i == 0:
                section_name = "Production"
                self.drill_collar_production = nearest_drill_collar
            elif i == total_iterations - 1:
                section_name = "Surface"
                self.drill_collar_surface = nearest_drill_collar
            else:
                section_name = "Intermediate"
                self.drill_collar_intermediate = nearest_drill_collar
            
            html_result += f"""
            <tr>
                <td>{section_name}</td>
                <td>{float(at_head):.1f}</td>
                <td>{bit_size:.2f}</td>
                <td>{nearest_drill_collar:.2f} mm</td>
            </tr>
            """
        
        html_result += "</table>"
        self.result_text.setHtml(html_result)

    def get_data_for_gamma(self, gamma_value):
        if self.df is None:
            return None
        try:
            gamma_value = float(gamma_value)
            matching_rows = self.df[np.isclose(self.df['γ'], gamma_value, atol=1e-8)]
            
            if matching_rows.empty:
                return None
            
            row = matching_rows.iloc[0]
            return {col: row[col] for col in self.additional_columns}
        
        except (ValueError, IndexError):
            return None

    def find_nearest(self, array, value):
        array = np.array(array)
        valid = ~np.isnan(array)
        return array[valid][np.argmin(np.abs(array[valid] - value))]

    def calculate_and_display(self):
        data = self.data_input_tab.get_data()
        calculation_html = "<h3>Results:</h3>"

        try:
            required_fields = ['WOB', 'C', 'qc', 'H', 'Lhw', 'qp', 'P', 'γ']
            for field in required_fields:
                for i in range(1, 4):
                    if not data.get(f"{field}_{i}"):
                        raise ValueError(f"Field '{field}' (Instance {i}) is empty")

            for i in range(1, 4):
                WOB = float(data[f'WOB_{i}'])
                C = float(data[f'C_{i}'])
                qc = float(data[f'qc_{i}'])
                H = float(data[f'H_{i}'])
                Lhw = float(data[f'Lhw_{i}'])
                qp = float(data[f'qp_{i}'])
                P = float(data[f'P_{i}'])
                γ = float(data[f'γ_{i}'])
                K1 = float(data['K1'])
                K2 = float(data['K2'])
                K3 = float(data['K3'])
                dα = float(data['dα'])
                Dep = float(data['Dep'])
                Dhw = float(data['Dhw'])
                n = float(data['n'])

                additional_data = self.get_data_for_gamma(γ)
                if additional_data:
                    b = additional_data.get('b', 0)
                    Mp = additional_data.get('Mp', 0)
                else:
                    b = float(data[f'b_{i}'])
                    Mp = float(data[f'Mp_{i}'])

                L0c = WOB / (C * qc * b)
                Lp = H - (Lhw + L0c)

                if additional_data:
                    Ap = additional_data.get('AP', 0)
                    Aip = additional_data.get('AIP', 0)
                    qhw = float(data['qhw'])

                    T = ((1.08 * Lp * qp + Lhw * qhw + L0c * qc) * b) / Ap
                    Tc = T + P * (Aip / Ap)
                    Tec = Tc * K1 * K2 * K3

                    if i == 1:
                        dec = self.drill_collar_production / 1000
                    elif i == 2:
                        dec = self.drill_collar_intermediate / 1000
                    else:
                        dec = self.drill_collar_surface / 1000

                    Np = dα * γ * (Lp * Dep**2 + L0c * dec**2 + Lhw * Dhw**2) * n**1.7

                    DB = self.nearest_bit_sizes[-i] / 1000
                    NB = 3.2 * 10**-2 * (WOB**0.5) * (DB**1.75) * n

                    tau = (30 * ((Np + NB) * 10**3 / (pi * n * Mp))) * 10**-6

                    eq = sqrt((Tec*10**-1)**2 + 4*tau**2)
                    C_new = eq * 1.5

                    nearest_mpi = self.find_nearest(self.drill_pipe_data['Minimum tensile strength(mpi)'], C_new)
                    metal_grade_index = self.drill_pipe_data['Minimum tensile strength(mpi)'].index(nearest_mpi)
                    metal_grade = self.drill_pipe_data['Drill pipe Metal grade'][metal_grade_index]

                    SegmaC = nearest_mpi

                    numerator = ((SegmaC/1.5)**2 - 4 * tau**2) * 10**12
                    denominator = ((7.85 - 1.5)**2) * 10**8
                    sqrt_result = sqrt(numerator / denominator)
                    Lmax = sqrt_result - ((L0c*qc + Lhw*qhw) / qp)

                    calculation_html += f"""
                    <h4>Instance {i}:</h4>
                    <p><strong>Drill pipe Metal grade:</strong> {metal_grade}</p>
                    <p><strong>Lmax:</strong> {Lmax:.2f}</p>
                    <br>
                    """

                else:
                    calculation_html += f"<p style='color: #F44747;'>No additional data found for the given γ value: {γ}</p>"

        except ValueError as e:
            calculation_html += f"<p style='color: #F44747;'>Error in calculations: {str(e)}</p>"
        except Exception as e:
            calculation_html += f"<p style='color: #F44747;'>Unexpected error: {str(e)}</p>"

        self.calculation_text.setHtml(calculation_html)
