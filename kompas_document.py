class KompasDocument:
    """Создает документ Компаса"""

    def __init__(self, doc):
        self.name = doc.Name
        self.path = doc.Path
        self.full_name = doc.PathName
        self.type = self.get_document_type(doc.DocumentType)
        self.__iDocument = doc
        self.iLayoutSheets = self.__iDocument.LayoutSheets
        self.library_path = r"e:\Program Files(x86)\ASCON\KOMPAS-3D V14\LYT\GRAPHICIL.LYT"

    @classmethod
    def get_document_type(cls, doc_type: int) -> str:
        types_dict = {0: 'Неизвестный тип',
                      1: 'Чертеж',
                      2: 'Фрагмент',
                      3: 'Спецификация',
                      4: 'Деталь',
                      5: 'Сборка',
                      6: 'Текстовый документ'}
        return types_dict[doc_type]

    @property
    def sheets_count(self):
        return self.iLayoutSheets.Count

    @property
    def iDocument(self):
        return self.__iDocument

    @staticmethod
    def comma_to_point(text):
        return text.replace(',', '.')

    @staticmethod
    def point_to_comma(text):
        return text.replace('.', ',')
