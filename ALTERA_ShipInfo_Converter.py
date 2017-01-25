#Job 8:7
#Though your beginning was small, yet your latter end would greatly increase.
import pandas as pd
import os, re
import codecs

class txt_coverter():
    def __init__(self):
        self.target_list = []
        self.shipment_list_count = 0
        self.lot_number_line = 0
        self.current_dir = os.getcwd()
        self.convert_df = pd.DataFrame(columns=['LOT#', 'P/D/L', 'Device', 'QTY', 'DATE CODE', 'TRACE CODE',])

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
                    data_combination = line.split()
                    # Point to reset all data
                    if line.find("PKG/DIM/LEAD..") != -1:
                        starting_point = line.find("PKG/DIM/LEAD..")
                        self.pdl = line[starting_point + 15:].replace(' ', '')
                        self.pdl = line[starting_point + 15:].rstrip()
                        self.lot_number_line = 0

                    elif line.find("DEVICE..") != -1:
                        starting_point = line.find("DEVICE..")
                        self.device_code = line[starting_point + 14:].replace(' ','')
                        self.device_code = line[starting_point + 14:].rstrip()

                    elif len(data_combination) == 5:
                        if data_combination[0] != 'ETD/FLIGHT':
                            print(data_combination)
                            self.lot_code = data_combination[0]
                            self.qty_code = data_combination[1]
                            self.trace_code = data_combination[3]
                        else:
                            pass

                    elif line.find("DATE CODE..") != -1:
                        starting_point = line.find("DATE CODE..")
                        self.date_code = line[starting_point + 11:].replace(' ','')
                        self.date_code = line[starting_point + 11:].rstrip()
                        self.convert_df.loc[len(self.convert_df)] = [self.lot_code, self.pdl, self.device_code, self.qty_code, self.date_code, self.trace_code, ]
                        self.lot_code = self.qty_code = self.date_code = self.trace_code = ""
                    else:
                        pass
                #'LOT#', 'Device', 'QTY', 'DATE CODE', 'TRACE CODE', 'E CODE', 'IC', 'COO'


                if not "ALTERA_ShipInfo_result" in os.listdir(self.current_dir):
                    os.mkdir("ALTERA_ShipInfo_result")

                file_quantity = len(os.listdir(os.path.join(self.current_dir,"ALTERA_ShipInfo_result")))
                writer = pd.ExcelWriter("{}/ALTERA_ShipInfo_result/{}-{}.xlsx".format(self.current_dir, file, file_quantity), engine="xlsxwriter")

                print(self.convert_df)

                self.convert_df.to_excel(writer, sheet_name="ALTERA")
                work_sheet = writer.sheets['ALTERA']
                work_sheet.set_column('A:A',5)
                work_sheet.set_column('B:B',20)
                work_sheet.set_column('C:C',10)
                work_sheet.set_column('D:D',21)
                work_sheet.set_column('E:E',8)
                work_sheet.set_column('F:F',15)
                work_sheet.set_column('G:G',12)
                writer.close()
                #reset for next file.
                self.convert_df = pd.DataFrame(columns=['LOT#', 'P/D/L', 'Device', 'QTY', 'DATE CODE', 'TRACE CODE', ])
                print("{} converting is done".format(file))

converter = txt_coverter()
converter.converting()