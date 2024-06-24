import xml.etree.ElementTree as ET
import pandas as pd
import csv
import math


def main():
    # Please replace file_path with your own XML and CSV file paths
    xml_extractor = XMLDataExtractor(file_path='G:\\Shared drives\\2023 Booker Engineering\\CLIENTS\\TRACK UTILITIES\\Butte\\O-Calc\\Resources\\Automation Test\\Midspan test\\10-4-Proposed.pplx')
    csv_extractor = CSVDataExtractor(file_path='G:\\Shared drives\\2023 Booker Engineering\\CLIENTS\\TRACK UTILITIES\\Butte\\O-Calc\\Resources\\Automation Test\\Midspan test\\10-4 data.csv')
    calculator = Calculator(xml_extractor, csv_extractor)
    results = calculator.calculate()


class XMLDataExtractor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None

    def read_data(self):
        tree = ET.parse(self.file_path)
        self.data = tree.getroot()

    def parse_data(self):
        spans = self.data.findall('.//Span')
        span_data_list = []
        for span in spans:
            attributes = span.find('ATTRIBUTES')
            span_data = {}
            for value in attributes.findall('VALUE'):
                if value.attrib['NAME'] in [
                    'SpanDistanceInInches',
                    'SpanEndHeightDelta',
                    'SpanType',
                    'Guid',
                    'Owner',
                    'ConductorDiameter'
                ]:
                    span_data[value.attrib['NAME']] = value.text
            if span_data:  # only add span data if it contains the desired attributes
                span_data_list.append(span_data)
        return span_data_list


class CSVDataExtractor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = []

    def read_data(self):
        with open(self.file_path, 'r') as f:
            reader = csv.reader(f)
            data = list(reader)

        # Find the line where "Element" is in the first column
        element_line_index = next((i for i, line in enumerate(data) if line and line[0] == 'Element'), None)

        if element_line_index is not None:
            # Read lines from the "Element" line onwards, stopping at the first blank line
            for line in data[element_line_index + 1:]:
                if not line or line[0].strip() == '':
                    break
                else:
                    self.data.append(line)

    def parse_data(self):
        # Create a DataFrame from the extracted lines
        from io import StringIO
        table_string = StringIO('\n'.join([','.join(line) for line in self.data]))
        df = pd.read_csv(table_string)

        # Filter the DataFrame to remove rows where 'SubType' is 'Overlashed Bundle' or 'Priority' is greater than '1'
        df = df.loc[~((df['SubType'] == 'Overlashed Bundle') | (df['Priority'] > 1))]

        cable_data = []
        for index, row in df.iterrows():
            cable_data.append({
                'AttachHeight': row['AttachHeight'],
                'CableWeightPerUnitLength': row['CableWeightPerUnitLength'],
                'CableDiameter': row['CableDiameter'],
                'CableLengthLead': row['CableLengthLead'],
                'SagMinimum': row['SagMinimum'],
                'SagNominal': row['SagNominal'],
                'SagMaximum': row['SagMaximum'],
            })

        # Return the list of dictionaries
        return cable_data


class EquationCalculator:
    def __init__(self, variable_dictionary):
        self.variable_dict = variable_dictionary

        # Calculate weight with ice and add to variables dict
        diameter = float(self.variable_dict['ConductorDiameter'])
        ice_weight = math.pi * ((diameter + 0.5)**2 - diameter**2) * 0.102
        conductor_weight = float(self.variable_dict['CableWeightPerUnitLength'])
        weight_with_ice = conductor_weight + ice_weight
        self.variable_dict['WeightWithIce'] = weight_with_ice

    def equation1_calculate_tension(self):
        # Pull variables
        w = float(self.variable_dict['CableWeightPerUnitLength'])
        wi = float(self.variable_dict['WeightWithIce'])
        s = float(self.variable_dict['SpanDistanceInInches']) / 12  # to feet
        sag_min = float(self.variable_dict['SagMinimum']) / 12  # to feet
        sag_nom = float(self.variable_dict['SagNominal']) / 12  # to feet
        sag_max = float(self.variable_dict['SagMaximum']) / 12  # to feet

        # Calc tensions
        t_min = (wi * s**2) / (8 * sag_min)
        t_nom = (w * s**2) / (8 * sag_nom)
        t_max = (w * s**2) / (8 * sag_max)

        # Add tensions to dictionary
        self.variable_dict['MinimumTension'] = t_min
        self.variable_dict['NominalTension'] = t_nom
        self.variable_dict['MaximumTension'] = t_max

    def equation2_calculate_x_r_values(self):
        # Pull variables
        w = float(self.variable_dict['CableWeightPerUnitLength'])
        wi = float(self.variable_dict['WeightWithIce'])
        s = float(self.variable_dict['SpanDistanceInInches']) / 12  # to feet
        h = float(self.variable_dict['SpanEndHeightDelta']) / 12  # to feet
        t_min = float(self.variable_dict['MinimumTension'])
        t_nom = float(self.variable_dict['NominalTension'])
        t_max = float(self.variable_dict['MaximumTension'])

        # Calc distance from pole
        x_r_min = (s / 2) - (t_min / wi) * math.asinh((abs(h) / 2) / ((t_min / wi) * math.sinh((s / 2) / (t_min / wi))))
        x_r_nom = (s / 2) - (t_nom / w) * math.asinh((abs(h) / 2) / ((t_nom / w) * math.sinh((s / 2) / (t_nom / w))))
        x_r_max = (s / 2) - (t_max / w) * math.asinh((abs(h) / 2) / ((t_max / w) * math.sinh((s / 2) / (t_max / w))))

        # Add distances to dictionary
        self.variable_dict['XrMinimum'] = x_r_min
        self.variable_dict['XrNominal'] = x_r_nom
        self.variable_dict['XrMaximum'] = x_r_max

    def equation3_change_in_height(self):
        # Pull variables
        w = float(self.variable_dict['CableWeightPerUnitLength'])
        wi = float(self.variable_dict['WeightWithIce'])
        t_min = float(self.variable_dict['MinimumTension'])
        t_nom = float(self.variable_dict['NominalTension'])
        t_max = float(self.variable_dict['MaximumTension'])
        x_r_min = float(self.variable_dict['XrMinimum'])
        x_r_nom = float(self.variable_dict['XrNominal'])
        x_r_max = float(self.variable_dict['XrMaximum'])

        # Calc change in height
        y_min = (wi * x_r_min**2) / (2 * t_min)
        y_nom = (w * x_r_nom**2) / (2 * t_nom)
        y_max = (w * x_r_max**2) / (2 * t_max)

        # Add change in height to dictionary
        self.variable_dict['YMinimum'] = y_min
        self.variable_dict['YNominal'] = y_nom
        self.variable_dict['YMaximum'] = y_max

    def equation4_calculate_lowest_point(self):
        # Pull variables
        h = float(self.variable_dict['SpanEndHeightDelta']) / 12  # to feet
        attach_height = float(self.variable_dict['AttachHeight'])
        y_min = float(self.variable_dict['YMinimum'])
        y_nom = float(self.variable_dict['YNominal'])
        y_max = float(self.variable_dict['YMaximum'])

        # Calc lowest points
        if h > 0:
            lp_min = attach_height - y_min
            lp_nom = attach_height - y_nom
            lp_max = attach_height - y_max
        else:
            lp_min = attach_height - abs(h) - y_min
            lp_nom = attach_height - abs(h) - y_nom
            lp_max = attach_height - abs(h) - y_max

        # Add to dictionary
        self.variable_dict['LPMinimum'] = lp_min
        self.variable_dict['LPNominal'] = lp_nom
        self.variable_dict['LPMaximum'] = lp_max

    def calculate_all(self):
        self.equation1_calculate_tension()
        self.equation2_calculate_x_r_values()
        self.equation3_change_in_height()
        self.equation4_calculate_lowest_point()

        return self.variable_dict


class Calculator:
    def __init__(self, xml_extractor, csv_extractor):
        self.xml_extractor = xml_extractor
        self.csv_extractor = csv_extractor

        self.xml_extractor.read_data()
        self.csv_extractor.read_data()

        xml_data = self.xml_extractor.parse_data()
        for span_data in xml_data:
            print(span_data)
        csv_data = self.csv_extractor.parse_data()
        for span_data in csv_data:
            print(span_data)

        if len(xml_data) != len(csv_data):
            raise ValueError("XML data and CSV data are of unequal length!")

        self.combined_data = [{**x, **y} for x, y in zip(xml_data, csv_data)]

    def calculate(self):
        calculated_results = []
        for data in self.combined_data:
            equation_calculator = EquationCalculator(data)
            result = equation_calculator.calculate_all()
            calculated_results.append(result)

        return calculated_results


if __name__ == '__main__':
    main()
