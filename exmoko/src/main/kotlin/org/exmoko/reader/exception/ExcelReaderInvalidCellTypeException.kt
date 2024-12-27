package org.exmoko.reader.exception

class ExcelReaderInvalidCellTypeException(
  message: String = "Invalid cell type."
) : RuntimeException(message)
