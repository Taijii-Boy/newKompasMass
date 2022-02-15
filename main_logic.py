from kompas_api import *

if __name__ == '__main__':
    kompas = KompasAPI()
    kompas.connect_to_kompas()
    api_document = kompas.get_active_doc()
    document = kompas.make_kompas_document(api_document)
    print(document.get_mass())


