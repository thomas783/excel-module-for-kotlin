package org.exmoko.writer.exception

abstract class ExcelWriterException(
  message: String = "An error occurred while creating an Excel file."
) : RuntimeException(message)
