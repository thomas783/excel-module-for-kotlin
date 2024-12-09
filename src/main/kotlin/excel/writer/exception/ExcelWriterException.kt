package excel.writer.exception

class ExcelWriterException(
  message: String = "An error occurred while creating an Excel file."
) : RuntimeException(message)
