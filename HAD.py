from PyQt5.QtWidgets import QWidget, QVBoxLayout, QTextEdit, QLabel
from PyQt5.QtGui import QIcon
import logging
import math

class HADCalculator(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setup_logging()
        self.depth = None

    def initUI(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.had_text = QTextEdit()
        self.had_text.setReadOnly(True)
        layout.addWidget(QLabel("HAD Results:"))
        layout.addWidget(self.had_text)

    def setup_logging(self):
        logging.basicConfig(level=logging.INFO, format='%(message)s')

    def update_had_results(self, had_data, depth, section_name):
        self.depth = depth
        self.had_text.clear()
        production_data = had_data.get(list(had_data.keys())[0], [])
        
        if production_data and section_name == "Production Section":
            sorted_data = sorted(production_data, key=lambda x: x['had'], reverse=True)
            
            if len(sorted_data) >= 3:
                l_values = self.calculate_l_values(sorted_data, depth)
                for i in range(min(3, len(sorted_data))):
                    sorted_data[i]['l_value'] = l_values[f'l{i+1}']
                if len(sorted_data) > 3 and 'l4' in l_values:
                    sorted_data[3]['l_value'] = l_values['l4']
            
            self._log_section_data(section_name, sorted_data)
            self.format_and_display_had_section(section_name, sorted_data)

    def _log_section_data(self, section_name, sorted_data):
        logging.info(f"\n{section_name}")
        logging.info("-" * 90)
        logging.info(f"{'Row':<5}{'HAD':<10}{'External Pressure':<20}{'Metal Type':<15}{'Tensile Strength':<20}{'Unit Weight':<15}{'L Value':<10}")
        logging.info("-" * 90)
        for i, data in enumerate(sorted_data, 1):
            l_value = data.get('l_value', '')
            l_value_str = f"{l_value:.2f}" if l_value != '' else ''
            logging.info(f"{i:<5}{data['had']:<10.2f}{data['external_pressure']:<20}{data['metal_type']:<15}{data['tensile_strength']:<20}{data['unit_weight']:<15}{l_value_str:<10}")
        logging.info("-" * 90)

    def format_and_display_had_section(self, section_name, data_list):
        section_html = self._generate_section_html(section_name, data_list)
        self.had_text.append(section_html)

    def _generate_section_html(self, section_name, data_list):
        header = self._generate_table_header(section_name)
        rows = self._generate_table_rows(data_list)
        return f"{header}{rows}</table></div>"

    def _generate_table_header(self, section_name):
        return f"""
        <div style='margin-top: 20px; margin-bottom: 20px;'>
            <h3 style='color: #4CAF50; padding: 10px;'>{section_name}</h3>
            <table style='border-collapse: collapse; width: 100%; background-color: #2b2b2b;'>
                <tr style='background-color: #4CAF50; color: white;'>
                    <th style='padding: 8px; text-align: center; border: 1px solid #555;'>Row</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #555;'>HAD</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #555;'>External Pressure (MPa)</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #555;'>Metal Type</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #555;'>Tensile Strength (Tonf)</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #555;'>Unit Weight (Lbs/ft)</th>
                    <th style='padding: 8px; text-align: center; border: 1px solid #555;'>L Value</th>
                </tr>
        """

    def _generate_table_rows(self, data_list):
        rows = ""
        for i, data in enumerate(data_list, 1):
            bg_color = "#3b3b3b" if i % 2 == 0 else "#2b2b2b"
            l_value = data.get('l_value', '')
            l_value_str = f"{l_value:.2f}" if l_value != '' else ''
            rows += f"""
                <tr style='background-color: {bg_color};'>
                    <td style='padding: 8px; text-align: center; color: white; border: 1px solid #555;'>{i}</td>
                    <td style='padding: 8px; text-align: right; color: white; border: 1px solid #555;'>{data['had']:.2f}</td>
                    <td style='padding: 8px; text-align: right; color: white; border: 1px solid #555;'>{data['external_pressure']}</td>
                    <td style='padding: 8px; text-align: center; color: white; border: 1px solid #555;'>{data['metal_type']}</td>
                    <td style='padding: 8px; text-align: right; color: white; border: 1px solid #555;'>{data['tensile_strength']}</td>
                    <td style='padding: 8px; text-align: right; color: white; border: 1px solid #555;'>{data['unit_weight']}</td>
                    <td style='padding: 8px; text-align: right; color: white; border: 1px solid #555;'>{l_value_str}</td>
                </tr>
            """
        return rows

    def calculate_l_values(self, data_list, depth):
        had_row_2 = data_list[1]['had']
        had_row_3 = data_list[2]['had']
        tensile_strength_row_2 = float(data_list[1]['tensile_strength'])
        tensile_strength_row_3 = float(data_list[2]['tensile_strength'])
        unit_weight_row_1 = float(data_list[0]['unit_weight'])
        unit_weight_row_2 = float(data_list[1]['unit_weight'])
        unit_weight_row_3 = float(data_list[2]['unit_weight'])

        best_l1 = self._find_best_l1(depth, had_row_2, tensile_strength_row_2, unit_weight_row_1)
        best_l2 = self._find_best_l2(depth, best_l1, had_row_3, tensile_strength_row_3, unit_weight_row_1, unit_weight_row_2)
        best_l3 = self._calculate_l3(best_l1, best_l2, tensile_strength_row_3, unit_weight_row_1, unit_weight_row_2, unit_weight_row_3)

        result = {'l1': best_l1, 'l2': best_l2, 'l3': best_l3}
        
        total_length = best_l1 + best_l2 + best_l3
        if total_length < depth and len(data_list) > 3:
            tensile_strength_row_4, unit_weight_row_4 = self._get_next_row_data(data_list)
            if tensile_strength_row_4 and unit_weight_row_4:
                best_l4 = self._calculate_l4(best_l1, best_l2, best_l3, tensile_strength_row_4, 
                                           unit_weight_row_1, unit_weight_row_2, unit_weight_row_3, unit_weight_row_4)
                result['l4'] = best_l4

        return result

    def _find_best_l1(self, depth, had_row_2, tensile_strength_row_2, unit_weight_row_1):
        def condition_difference(l1):
            y1, z1 = self._calculate_y1_z1(l1, depth, had_row_2, tensile_strength_row_2, unit_weight_row_1)
            return abs(y1**2 + z1**2 + y1*z1 - 1.00)

        best_l1 = 0
        min_difference = float('inf')
        for i in range(1, int(depth)):
            diff = condition_difference(i)
            if diff < min_difference:
                min_difference = diff
                best_l1 = i
            if diff < 0.0001:
                break
        return best_l1

    def _calculate_y1_z1(self, l1, depth, had_row_2, tensile_strength_row_2, unit_weight_row_1):
        y1 = (depth - l1) / had_row_2
        z1 = (l1 * unit_weight_row_1 * 1.488) / (tensile_strength_row_2 * 1000)
        return y1, z1

    def _find_best_l2(self, depth, l1, had_row_3, tensile_strength_row_3, unit_weight_row_1, unit_weight_row_2):
        def condition_difference(l2):
            y2, z2 = self._calculate_y2_z2(l1, l2, depth, had_row_3, tensile_strength_row_3, unit_weight_row_1, unit_weight_row_2)
            return abs(y2**2 + z2**2 + y2*z2 - 1.00)

        best_l2 = 0
        min_difference = float('inf')
        for i in range(1, int(depth - l1)):
            diff = condition_difference(i)
            if diff < min_difference:
                min_difference = diff
                best_l2 = i
            if diff < 0.0001:
                break
        return best_l2

    def _calculate_y2_z2(self, l1, l2, depth, had_row_3, tensile_strength_row_3, unit_weight_row_1, unit_weight_row_2):
        y2 = (depth - (l1 + l2)) / had_row_3
        z2 = (l2 * unit_weight_row_1 + l2 * unit_weight_row_2) * 1.488 / (tensile_strength_row_3 * 1000)
        return y2, z2

    def _calculate_l3(self, l1, l2, tensile_strength_row_3, unit_weight_row_1, unit_weight_row_2, unit_weight_row_3):
        return ((tensile_strength_row_3 * 1000 / 1.75) - (l1 * unit_weight_row_1 * 1.488 + l2 * unit_weight_row_2 * 1.488)) / (unit_weight_row_3 * 1.488)

    def _get_next_row_data(self, data_list):
        if len(data_list) > 3:
            return float(data_list[3]['tensile_strength']), float(data_list[3]['unit_weight'])
        return None, None

    def _calculate_l4(self, l1, l2, l3, tensile_strength_row_4, unit_weight_row_1, unit_weight_row_2, unit_weight_row_3, unit_weight_row_4):
        return ((tensile_strength_row_4 * 1000 / 1.75) - (l1 * unit_weight_row_1 * 1.488 + l2 * unit_weight_row_2 * 1.488 + l3 * unit_weight_row_3 * 1.488)) / (unit_weight_row_4 * 1.488)

    def get_next_row_data(self, file_path, at_head_value, metal_type):
        db_calculator = self.parent().parent().db_calculator
        return db_calculator.get_next_row_data(file_path, at_head_value, metal_type)
##comment