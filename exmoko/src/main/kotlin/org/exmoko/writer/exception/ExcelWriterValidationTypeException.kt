package org.exmoko.writer.exception

abstract class ExcelWriterValidationTypeException(
  message: String = "ExcelColumn with validationType is something wrong..."
) : ExcelWriterException(message)

class ExcelWriterValidationListException(
  message: String = "ExcelColumn with either validationListOptions or validationListEnum is required"
) : ExcelWriterValidationTypeException(message)

class ExcelWriterValidationDecimalException(
  message: String = "ExcelColumn with operationType is required"
) : ExcelWriterValidationTypeException(message)

class ExcelWriterValidationIntegerException(
  message: String = "ExcelColumn with operationType is required"
) : ExcelWriterValidationTypeException(message)

class ExcelWriterValidationFormulaException(
  message: String = "ExcelColumn with validationFormula is required",
) : ExcelWriterValidationTypeException(message)

class ExcelWriterValidationTextLengthException(
  message: String = "ExcelColumn with operationType is required"
) : ExcelWriterValidationTypeException(message)
