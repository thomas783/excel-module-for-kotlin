package excel.reader.exception

class ExcelReaderException(
  message: String = "An error occurred while reading an Excel file."
) : RuntimeException(message)
