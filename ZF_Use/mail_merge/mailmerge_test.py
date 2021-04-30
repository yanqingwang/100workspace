from mailmerge import MailMerge
import time

from win32com.client import gencache, DispatchEx

class writedocx():
    def __init__(self, dict, filepath, savepath):
        super().__init__()
        self.filepath = filepath
        self.savepath = savepath
        self.dict = dict
        self.writedoc(self.dict)

    def writedoc(self,dict):
        document_1 = MailMerge(self.filepath)
        print (document_1.get_merge_fields())
        dict = self.dict
        document_1.merge(** dict)
        document_1.write(self.savepath)

    def word2pdf(self,pdf_file):
        wdFormatPDF = 17
        # w = DispatchEx("Word.Application")
        w = comtypes.client.CreateObject('Word.Application')
        w.Visible = False
        try:
            doc = w.Documents.Open(self.savepath, ReadOnly=1)
            # doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
            doc.ExportAsFixedFormat(OutputFileName=pdf_file,
                                    ExportFormat=17,  # 17 = PDF output, 18=XPS output
                                    OpenAfterExport=False,
                                    OptimizeFor=0,  # 0=Print (higher res), 1=Screen (lower res)
                                    CreateBookmarks=1,
                                    # 0=No bookmarks, 1=Heading bookmarks only, 2=bookmarks match word bookmarks
                                    DocStructureTags=True
                                    );
            doc.Close()
        finally:
            w.Quit()


if __name__ == '__main__':
    time1 = time.time()

    d_range = {"姓名":'zhangsan',"年龄":"20","生日":"1988-08-08","folder":'folder1'}
    r_writer = writedocx(d_range,'C:/temp/mm/Test.docx','C:/temp/mm/Mail/test01.docx')
    r_writer.word2pdf('C:/temp/mm/Mail/test01.pdf')

    print("Done, Total running time", time.time() - time1)
