package excel.writer.exception

class ExcelWriterValidationListException(
  message: String = "ExcelColumn with either validationListOptions or validationListEnum is required"
) : ExcelWriterException(message)
