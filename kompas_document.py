class KompasDocument:
    """Создает документ Компаса"""

    def __init__(self, doc):
        self.name = doc.Name
        self.path = doc.Path
        self.full_name = doc.PathName
        self.type = self.get_document_type(doc.DocumentType)
        self.api = doc

    @staticmethod
    def get_document_type(type: int) -> str:
        types_dict = {0: 'Неизвестный тип',
                      1: 'Чертеж',
                      2: 'Фрагмент',
                      3: 'Спецификация',
                      4: 'Деталь',
                      5: 'Сборка',
                      6: 'Текстовый документ'}
        return types_dict[type]

    def get_views(self):
        pass