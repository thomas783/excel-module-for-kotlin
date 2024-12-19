package excel.writer.exception

class ExcelWritableMissingException(
  message: String = "ExcelWritable annotation is required"
) : ExcelWriterException(message)
