package writer.dto

import shared.IExcelWriterCommonDto

data class ExcelWriterWritableMissingErrorDto(
  val id: Long,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterWritableMissingErrorDto> {
    override fun createSampleData(size: Int): Collection<ExcelWriterWritableMissingErrorDto> {
      return (1..size).map { number ->
        ExcelWriterWritableMissingErrorDto(
          id = number.toLong()
        )
      }
    }
  }
}
