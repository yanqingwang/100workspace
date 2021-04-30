from mailmerge import MailMerge
import time
import pandas as pd
import os
from docx2pdf import convert
# from win32com.client import DispatchEx


class handle_merge():
    def __init__(self):
        self.df_data = pd.DataFrame()
        self.template = 'C:/temp/mm/Test.docx'
        self.datafile = 'C:/temp/mm/Test.xlsx'
        self.path = 'C:/temp/mm/'
        self.document_1 = MailMerge(self.template)

    def read_files(self,):
        # df_data = pd.DataFrame()
        try:
            self.df_data = pd.DataFrame(pd.read_excel(io=self.datafile, sheet_name='Sheet1',
                                                       # dtype={'ZF Global ID': 'str'},
                                                       header=0,
                                                       # skiprows=2))
                                                       ))
            print(self.df_data.head())
        except Exception as e:
            print('Handle Employee Data Exception:', self.datafile, e)


    # def generate_pdf(self,docfile, pdf_file):
    #     # wdFormatPDF = 17
    #     word = DispatchEx("Word.Application")
    #     try:
    #         doc = word.Documents.Open(docfile, ReadOnly=1)
    #         # doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
    #         doc.ExportAsFixedFormat(OutputFileName=pdf_file,
    #                                 ExportFormat=17,  # 17 = PDF output, 18=XPS output
    #                                 OpenAfterExport=False,
    #                                 OptimizeFor=0,  # 0=Print (higher res), 1=Screen (lower res)
    #                                 CreateBookmarks=1,
    #                                 # 0=No bookmarks, 1=Heading bookmarks only, 2=bookmarks match word bookmarks
    #                                 DocStructureTags=True
    #                                 );
    #         doc.Close()
    #     except Exception as e:
    #         print('convert pdf failed,', docfile)


    def main(self):
        print (self.document_1.get_merge_fields())
        self.read_files()

        for n_idx,n_row in self.df_data.iterrows():
            dict = {
                "姓名": str(n_row['姓名']),
                "年龄": str(n_row['年龄']),
                "生日": str(n_row['生日']),
                "folder": 'folder1'
            }
            self.document_1.merge(** dict)

            file_path = self.path + 'xlsx/'
            if not os.path.exists(file_path):
                os.makedirs(file_path)

            file_docx = self.path + 'xlsx/' + n_row['姓名'] + '.docx'
            self.document_1.write(file_docx)
            # time.sleep(1)

            file_path = self.path + n_row['性别'] + '/'
            if not os.path.exists(file_path):
                os.makedirs(file_path)

            file_pdf = file_path + n_row['姓名'] + '.pdf'
            # file_pdf = self.path + 'pdfs/' + n_row['姓名'] + '.pdf'
            # self.generate_pdf(file_docx,file_pdf)
            convert(file_docx,file_pdf,keep_active=True)
        self.document_1.close()



if __name__ == '__main__':

    time1 = time.time()
    mail_merge = handle_merge()
    mail_merge.main()

    print("Done, Total running time", time.time() - time1)
