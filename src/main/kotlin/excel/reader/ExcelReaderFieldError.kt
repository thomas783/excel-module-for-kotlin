package excel.reader

enum class ExcelReaderFieldError(var message: String) {
  HEADER_MISSING("Essential header missing"), TYPE("Invalid data type: "), VALID("Validation error"), UNKNOWN("Unknown");

  companion object {
    private var messageToMap: MutableMap<String, ExcelReaderFieldError> = mutableMapOf()

    fun getExcelReaderErrorConstant(name: String): ExcelReaderFieldError? {
      if (messageToMap.isEmpty()) {
        initMapping()
      }
      return messageToMap[name]
    }

    private fun initMapping() {
      messageToMap = mutableMapOf()
      values().forEach {
        messageToMap[it.name] = it
      }
    }
  }
}
