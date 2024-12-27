package org.exmoko.reader.exception

class ExcelReaderFileExtensionException(
  message: String = "ExcelReader only supports for .xlsx and .xls file extensions"
) : ExcelReaderException(message)
