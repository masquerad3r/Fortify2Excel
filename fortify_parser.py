import xml.etree.ElementTree as ET
import xlsxwriter
import sys

try:
    #  Taking input
    xml_file_name = sys.argv[1]

    #  Checking for xml extension
    if xml_file_name.split(".")[-1] != "xml":
        print("Please provide an xml file only")
        exit()

    #  Parsing the xml file
    tree = ET.parse(xml_file_name)
    root = tree.getroot()


#  Handling absence of files in command lines
except IndexError as e:
    print("Usage: python fortify_parser.py <xml file name>")
    exit()

#  Handling improper structured xml file
except ET.ParseError as e:
    print("Could not parse the given XML file, please check the format")
    exit()

#  Variable Declaration
report = {}
counter = 1

compiled_info = {'critical': [],
                 'high': [],
                 'medium': [],
                 'low': []}

report_root = root.findall(".//ReportSection[3]/SubSection/IssueListing/Chart//GroupingSection")


def main():
    #  Creating a new Excel file
    workbook = xlsxwriter.Workbook('Fortify Report.xlsx')

    #  Worksheets declaration
    worksheet = workbook.add_worksheet('Static Scan Results')

    #  Initialising text formatting dictionary
    text_format = set_text_format(workbook)

    #  Setting the header fields
    set_headers(worksheet, text_format['header'])

    for parent in report_root:

        report["security_risk"] = parent.find("groupTitle").text
        child = parent.findall('Issue')

        for grand_child in child:

            #  Fetching header objects
            severity = grand_child.find("Folder")
            description = grand_child.find("Abstract")
            source_file_name = grand_child.find("Source/FileName")
            source_file_path = grand_child.find("Source/FilePath")
            source_line_number = grand_child.find("Source/LineStart")
            sink_file_name = grand_child.find("Primary/FileName")
            sink_file_path = grand_child.find("Primary/FilePath")
            sink_line_number = grand_child.find("Primary/LineStart")
            comments = grand_child.findall("Comment")

            #  Preparing the dictionary of report
            if severity is not None:
                report["severity"] = severity.text
            else:
                report["severity"] = ''

            if description is not None:
                report["description"] = description.text
            else:
                report["description"] = ''

            if source_file_name is not None:
                report["source_file_name"] = source_file_name.text
            else:
                report["source_file_name"] = ''

            if source_file_path is not None:
                report["source_file_path"] = source_file_path.text
            else:
                report["source_file_path"] = ''

            if source_line_number is not None:
                report["source_line_number"] = source_line_number.text
            else:
                report["source_line_number"] = ''

            if sink_file_name is not None:
                report["sink_file_name"] = sink_file_name.text
            else:
                report["sink_file_name"] = ''

            if sink_file_path is not None:
                report["sink_file_path"] = sink_file_path.text
            else:
                report["sink_file_path"] = ''

            if sink_line_number is not None:
                report["sink_line_number"] = sink_line_number.text
            else:
                report["sink_line_number"] = ''

            #  Extracting comments
            actual_comment = ''

            if comments is not None:
                for sub_comment in comments:
                    comment = sub_comment.find('Comment')
                    actual_comment += comment.text + "\n"

                report['comments'] = actual_comment[:-1]  # For removing the extra new line character

            #  Compile the results
            compile_report(report)

    #  Populate the values in excel file
    print_report(worksheet, text_format['severity'], text_format['normal'])

    #  Setting the zoom factor of the excel sheet
    worksheet.set_zoom(70)

    #  Closing the workbook
    workbook.close()


#  Function to compile report into their respective categories
def compile_report(report):
    if bool(report['severity']):

        if report['severity'] == 'Critical':
            compiled_info['critical'].append((report['security_risk'], report['severity'], report['description'],
                                              report['source_file_name'], report['source_file_path'],
                                              report['source_line_number'], report['sink_file_name'],
                                              report['sink_file_path'], report['sink_line_number'], report['comments']))

        if report['severity'] == 'High':
            compiled_info['high'].append((report['security_risk'], report['severity'], report['description'],
                                          report['source_file_name'], report['source_file_path'],
                                          report['source_line_number'], report['sink_file_name'], report['sink_file_path'],
                                          report['sink_line_number'], report['comments']))

        if report['severity'] == 'Medium':
            compiled_info['medium'].append((report['security_risk'], report['severity'], report['description'],
                                            report['source_file_name'], report['source_file_path'],
                                            report['source_line_number'], report['sink_file_name'], report['sink_file_path'],
                                            report['sink_line_number'], report['comments']))

        if report['severity'] == 'Low':
            compiled_info['low'].append((report['security_risk'], report['severity'], report['description'],
                                         report['source_file_name'], report['source_file_path'], report['source_line_number'],
                                         report['sink_file_name'], report['sink_file_path'], report['sink_line_number'],
                                         report['comments']))


#  Function to set header fields
def set_headers(worksheet, header_format):
    #  Setting the column width
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 40)
    worksheet.set_column('C:C', 10)
    worksheet.set_column('D:D', 50)
    worksheet.set_column('E:E', 20)
    worksheet.set_column('F:F', 40)
    worksheet.set_column('G:G', 11)
    worksheet.set_column('H:H', 20)
    worksheet.set_column('I:I', 40)
    worksheet.set_column('J:J', 11)
    worksheet.set_column('K:K', 20)

    #  Populating the header fields
    #  Serial Number
    worksheet.write("A1", "S.No.", header_format)

    #  Security Risk
    worksheet.write("B1", "Security Risk", header_format)

    #  Severity
    worksheet.write("C1", "Severity", header_format)

    #  Description
    worksheet.write("D1", "Description", header_format)

    #  Source File Name
    worksheet.write("E1", "Source File Name", header_format)

    #  Source File Path
    worksheet.write("F1", "Source File Path", header_format)

    #  Source Line number
    worksheet.write("G1", "Line Number", header_format)

    #  Sink File Name
    worksheet.write("H1", "Sink File Name", header_format)

    #  Sink File Path
    worksheet.write("I1", "Sink File Path", header_format)

    #  Sink Line Number
    worksheet.write("J1", "Line Number", header_format)

    #  Remarks
    worksheet.write("K1", "Remarks", header_format)


#  Function to set text formatting
def set_text_format(workbook):
    #  Setting the header text formatting
    header_format = workbook.add_format({
        'bold': True,
        'border': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#1E88E5',  # Blue
        'font_color': '#FFFFFF'})  # White

    #  Setting the text formatting
    normal_text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'border': True,
        'text_wrap': True})

    #  Setting formatting for severity levels
    #  Critical
    critical_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'font_color': '#FFFFFF',  # White
        'bg_color': '#FF0000'})  # Red

    #  High
    high_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'font_color': '#FFFFFF',  # White
        'bg_color': '#FF6E00'})  # Orange

    #  Medium
    medium_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'font_color': '#FFFFFF',  # White
        'bg_color': '#FFBB00'})  # Yellow

    #  Low
    low_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': True,
        'font_color': '#FFFFFF',  # White
        'bg_color': '#27AE60'})  # Green

    text_format = {'header': header_format,
                   'normal': normal_text_format,
                   'severity': {
                       'Critical': critical_format,
                       'High': high_format,
                       'Medium': medium_format,
                       'Low': low_format}
                   }

    return text_format


#  Function to sort the report by severity
def print_report(worksheet, severity_format, text_format):
    #  Variable declaration
    global counter

    #  Writing information according to severity of the bug
    for sequence in ("critical", "high", "medium", "low"):
        for i in compiled_info[sequence]:
            row = counter + 1

            #  Setting Serial number
            worksheet.write(f'A{row}', counter, text_format)

            #  Setting security risk
            worksheet.write(f'B{row}', i[0], text_format)

            #  Setting severity
            worksheet.write(f'C{row}', i[1], severity_format[i[1]])

            #  Setting Description
            worksheet.write(f'D{row}', i[2], text_format)

            #  Setting Source File Name
            worksheet.write(f'E{row}', i[3], text_format)

            #  Setting Source File Path
            worksheet.write(f'F{row}', i[4], text_format)

            #  Setting Source Line Number
            worksheet.write(f'G{row}', i[5], text_format)

            #  Setting Sink File Name
            worksheet.write(f'H{row}', i[6], text_format)

            #  Setting Sink File Path
            worksheet.write(f'I{row}', i[7], text_format)

            #  Setting Sink Line Number
            worksheet.write(f'J{row}', i[8], text_format)

            #  Setting Comments
            worksheet.write(f'K{row}', i[9], text_format)

            counter += 1


if __name__ == "__main__":
    main()

    print("Report created successfully")
