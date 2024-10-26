import json
import os
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
                             QPushButton, QGroupBox, QGridLayout, QScrollArea, QSpacerItem,
                             QSizePolicy, QToolTip)
from PyQt5.QtGui import QFont, QColor, QPalette, QIcon
from PyQt5.QtCore import Qt, QSize

class DataInputTab(QWidget):
    def __init__(self):
        super().__init__()
        self.data_file = 'saved_data.json'
        self.initUI()
        self.load_saved_data()

    def initUI(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(20)

        groups = [
            ("Basic Parameters", ["WOB", "C", "qc", "qp", "Lhw", "P", "γ", "H"]),
            ("Well Specifications", ["Dep", "Dhw", "qhw", "dα", "n"]),
            ("Constants", ["K1", "K2", "K3"])
        ]

        for title, fields in groups:
            group = self.create_group(title, fields)
            scroll_layout.addWidget(group)

        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)

        button_layout = QHBoxLayout()
        save_button = QPushButton("Save Data")
        save_button.setIcon(QIcon.fromTheme("document-save"))
        save_button.setIconSize(QSize(24, 24))
        save_button.setStyleSheet(self.get_button_style())
        save_button.clicked.connect(self.save_data)
        button_layout.addStretch()
        button_layout.addWidget(save_button)

        main_layout.addLayout(button_layout)

        self.setStyleSheet(self.get_dark_theme_style())

    def create_group(self, title, fields):
        group = QGroupBox(title)
        group.setStyleSheet(self.get_group_style())
        layout = QGridLayout()
        layout.setColumnStretch(1, 1)
        layout.setVerticalSpacing(10)
        layout.setHorizontalSpacing(15)

        for i, field in enumerate(fields):
            label = QLabel(field)
            label.setFont(QFont("Roboto", 10))
            layout.addWidget(label, i, 0)

            if title == "Basic Parameters":
                for j in range(3):
                    input_field = QLineEdit()
                    input_field.setFont(QFont("Roboto", 10))
                    input_field.setStyleSheet(self.get_input_style())
                    input_field.setPlaceholderText(f"Enter {field} ({j+1})")
                    tooltip_text = f"Enter the value for {field} (Instance {j+1})"
                    input_field.setToolTip(tooltip_text)
                    layout.addWidget(input_field, i, j+1)
                    setattr(self, f"{field}_{j+1}", input_field)
            else:
                input_field = QLineEdit()
                input_field.setFont(QFont("Roboto", 10))
                input_field.setStyleSheet(self.get_input_style())
                input_field.setPlaceholderText(f"Enter {field}")
                tooltip_text = f"Enter the value for {field}"
                input_field.setToolTip(tooltip_text)
                layout.addWidget(input_field, i, 1)
                setattr(self, field, input_field)

        group.setLayout(layout)
        return group

    def save_data(self):
        data = self.get_data()
        with open(self.data_file, 'w') as f:
            json.dump(data, f)
        print("Data saved!")

    def load_saved_data(self):
        if os.path.exists(self.data_file):
            with open(self.data_file, 'r') as f:
                data = json.load(f)
            self.set_data(data)

    def set_data(self, data):
        for field, value in data.items():
            if hasattr(self, field):
                getattr(self, field).setText(str(value))

    def get_data(self):
        data = {}
        for field in ["WOB", "C", "qc", "qp", "Lhw", "P", "γ", "H"]:
            for i in range(1, 4):
                data[f"{field}_{i}"] = getattr(self, f"{field}_{i}").text()
        
        for field in ["K1", "K2", "K3", "Dep", "Dhw", "qhw", "dα", "n"]:
            data[field] = getattr(self, field).text()
        
        return data

    @staticmethod
    def get_dark_theme_style():
        return """
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QLabel {
                color: #ffffff;
            }
            QToolTip {
                background-color: #3b3b3b;
                color: #ffffff;
                border: 1px solid #555555;
                padding: 5px;
            }
        """

    @staticmethod
    def get_group_style():
        return """
            QGroupBox {
                font-weight: bold;
                border: 2px solid #4CAF50;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 10px;
                color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                background-color: #2b2b2b;
            }
        """

    @staticmethod
    def get_input_style():
        return """
            QLineEdit {
                padding: 8px;
                border: 1px solid #555555;
                border-radius: 5px;
                background-color: #3b3b3b;
                color: #ffffff;
            }
            QLineEdit:focus {
                border: 2px solid #4CAF50;
            }
            QLineEdit::placeholder {
                color: #888888;
            }
        """

    @staticmethod
    def get_button_style():
        return """
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px 20px;
                font-size: 16px;
                font-weight: bold;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """
