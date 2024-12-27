package org.excelkotlin.reader

class ExcelReaderErrorField(
  var type: String?,
  var row: Int?,
  var field: String?,
  var fieldHeader: String?,
  var inputData: String?,
  var message: String?,
  var exceptionMessage: String?,
) {
  constructor() : this(
    type = null,
    row = null,
    field = null,
    fieldHeader = null,
    inputData = null,
    message = null,
    exceptionMessage = null
  )

  override fun toString(): String {
    return "Error Type=$type, Row=$row, Target Field Name=$field, Input Field Name=$fieldHeader, Input Value=$inputData, Error Type=$message, Error Message=$exceptionMessage"
  }
}
