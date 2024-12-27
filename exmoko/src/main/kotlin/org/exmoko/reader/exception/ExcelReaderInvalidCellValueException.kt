package org.exmoko.reader.exception

class ExcelReaderInvalidCellValueException(
  message: String = "Invalid cell value."
) : RuntimeException(message)
