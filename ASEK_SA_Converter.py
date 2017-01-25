#Job 8:7
#Though your beginning was small, yet your latter end would greatly increase.

import pandas as pd
import os, re
import codecs

class txt_coverter():
    def __init__(self):
        self.target_list = []
        self.e_code_finder = 0
        self.packing_list_count = 0
        self.current_dir = os.getcwd()
        self.convert_df = pd.DataFrame(columns=['LOT#', 'Device', 'QTY', 'DATE CODE', 'TRACE CODE', 'E CODE', 'IC', 'COO' ])

        txt_finder = re.compile(r'.*[.]txt', re.IGNORECASE)
        for file in os.listdir(self.current_dir):
            target_file = txt_finder.findall(file)
            if target_file != []:
                self.target_list.append(file)

    def converting(self):
        for idex, file in enumerate(self.target_list):
            # f = pd.read_clipboard()
            with codecs.open(file, 'r', encoding="UTF-8", errors='ignore') as f:
                lines = f.readlines()
                for line in lines:
                    if line.find("ASE SHIP ALERT REPORT") != -1:
                        self.lot_code = self.device_code = self.qty_code = self.date_code = self.trace_code = self.e_code = self.ic_code = self.coo = ""
                        self.packing_list_count = 0

                    elif line.find("PACKING   LIST") != -1:
                        if self.packing_list_count == 0:
                            self.convert_df.loc[len(self.convert_df)] = [self.lot_code, self.device_code, self.qty_code, self.date_code, self.trace_code, self.e_code, self.ic_code, self.coo]
                            self.packing_list_count += 1
                        else:
                            pass

                    elif line.find(" DATE CODE :") != -1:
                        starting_point = line.find(" DATE CODE :")
                        self.date_code = line[starting_point + 12:].replace(' ','')
                        self.date_code = line[starting_point + 12:].rstrip()

                    elif line.find(" TRACE CODE:") != -1:
                        starting_point = line.find(" TRACE CODE:")
                        self.trace_code = line[starting_point+12:]
                        self.trace_code = line[starting_point + 12:].rstrip()

                    elif line.find(" E CODE    :")  != -1:
                        self.e_code_finder = 1

                    elif self.e_code_finder == 1:
                        starting_point = line.find("14S/OPN info.")
                        self.e_code = line[starting_point+14:].replace(' ','')
                        self.e_code = line[starting_point + 14:].rstrip()
                        self.e_code_finder = 0

                    elif line.find("  IC") != -1:
                        starting_point = line.find("  IC")
                        self.ic_code = line[starting_point+5:]
                        self.ic_code = line[starting_point + 5:].rstrip()

                    elif line.find("  DEV#") != -1:
                        starting_point = line.find("  DEV#")
                        self.device_qty = line[starting_point+6:]
                        self.device_qty = self.device_qty.split()
                        self.device_code = self.device_qty[0]
                        self.qty_code = self.device_qty[1]

                    elif line.find("MADE IN") != -1:
                        starting_point = line.find("MADE IN")
                        self.coo = line[starting_point+8:].replace(' ','')
                        self.coo = line[starting_point + 8:].rstrip()

                    elif line.find("  CUST LOT:") != -1:
                        starting_point = line.find("  CUST LOT:")
                        self.lot_code = line[starting_point+11:].replace(' ','')
                        self.lot_code = line[starting_point + 11:].rstrip()

                    else:
                        pass
                #'LOT#', 'Device', 'QTY', 'DATE CODE', 'TRACE CODE', 'E CODE', 'IC', 'COO'


                if not "ASEK_SA_result" in os.listdir(self.current_dir):
                    os.mkdir("ASEK_SA_result")

                file_quantity = len(os.listdir(os.path.join(self.current_dir,"ASEK_SA_result")))
                writer = pd.ExcelWriter("{}/ASEK_SA_result/{}-{}.xlsx".format(self.current_dir, file, file_quantity), engine="xlsxwriter")

                self.convert_df.to_excel(writer, sheet_name="ASEK")
                work_sheet = writer.sheets['ASEK']
                work_sheet.set_column('B:B',15)
                work_sheet.set_column('C:C',26)
                work_sheet.set_column('D:D',5)
                work_sheet.set_column('E:E',14)
                work_sheet.set_column('F:F',13)
                work_sheet.set_column('G:G',21)
                work_sheet.set_column('H:H',34)
                work_sheet.set_column('I:I',8)
                writer.close()
                #reset for next file.
                self.convert_df = pd.DataFrame(columns=['LOT#', 'Device', 'QTY', 'DATE CODE', 'TRACE CODE', 'E CODE', 'IC', 'COO'])
                print("{} converting is done".format(file))


converter = txt_coverter()
converter.converting()

