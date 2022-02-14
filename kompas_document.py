class KompasDocument:
    """Создает документ Компаса"""

    def __init__(self, doc):
        self.name = doc.Name
        self.path = doc.Path
        self.full_name = doc.PathName
        self.type = self.get_document_type(doc.DocumentType)
        self.__iDocument = doc
        self.__iLayoutSheets = self.iDocument.LayoutSheets

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
        return self.__iLayoutSheets.Count

    @property
    def iDocument(self):
        return self.__iDocument

    def get_views(self):
        pass


class CDW(KompasDocument):
    def __init__(self, doc):
        super().__init__(doc)


class SPW(KompasDocument):
    pass
