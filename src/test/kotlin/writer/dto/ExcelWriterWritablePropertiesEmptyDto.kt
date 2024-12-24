package writer.dto

import com.excelkotlin.writer.annotation.ExcelWritable
import com.excelkotlin.writer.annotation.ExcelWriterHeader
import shared.IExcelWriterCommonDto

@ExcelWritable
data class ExcelWriterWritablePropertiesEmptyDto(
  @ExcelWriterHeader(name = "FIRST")
  val first: String,
  val second: String,
  val third: String,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterWritablePropertiesEmptyDto> {
    override fun createSampleData(size: Int): Collection<ExcelWriterWritablePropertiesEmptyDto> {
      return (1..size).map { number ->
        ExcelWriterWritablePropertiesEmptyDto(
          first = "first $number",
          second = "second $number",
          third = "third $number",
        )
      }
    }
  }
}
