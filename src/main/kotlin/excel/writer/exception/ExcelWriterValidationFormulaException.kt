package excel.writer.exception

class ExcelWriterValidationFormulaException(
  message: String = "ExcelColumn with validationFormula is required",
) : ExcelWriterException(message)
