from enum import Enum


# ENUM для представления констант всех доступных расширений файлов
class FileExtensions(Enum):
    TXT = "txt"
    XML = "xml"
    DOCX = "docx"
    XLSX = "xlsx"
    PDF = "pdf"

    # Метод для получения значений всех возможных расширений файлов
    @classmethod
    def as_list(cls):
        return [cls.TXT.value, cls.PDF.value, cls.XLSX.value, cls.DOCX.value, cls.XML.value]
