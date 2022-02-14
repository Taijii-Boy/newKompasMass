from kompas_api import *

if __name__ == '__main__':
    kompas = KompasAPI()
    kompas.connect_to_kompas()
    doc = kompas.get_active_doc()
    print(doc.sheets_count)
    print(doc.iDocument)
    cdw = CDW(doc)
    print(cdw)
