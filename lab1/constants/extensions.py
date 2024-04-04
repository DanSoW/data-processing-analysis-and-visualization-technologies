from enum import Enum


# Enum for defined all accessible file extension
class FileExtensions(Enum):
    TXT = "txt"
    XML = "xml"
    DOCX = "docx"
    XLSX = "xlsx"
    PDF = "pdf"

    @classmethod
    def as_list(cls):
        return [cls.TXT.value, cls.PDF.value, cls.XLSX.value, cls.DOCX.value, cls.XML.value]
