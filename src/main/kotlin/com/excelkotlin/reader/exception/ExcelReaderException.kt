package com.excelkotlin.reader.exception

open class ExcelReaderException(
  message: String = "An error occurred while reading an Excel file."
) : RuntimeException(message)