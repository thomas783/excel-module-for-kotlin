package org.exmoko.writer.exception

class ExcelWritableMissingException(
  message: String = "ExcelWritable annotation is required"
) : ExcelWriterException(message)
