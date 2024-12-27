package writer.dto.validation

import org.excelkotlin.writer.annotation.ExcelWritable
import org.excelkotlin.writer.annotation.ExcelWriterColumn
import org.excelkotlin.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import writer.dto.IExcelWriterCommonDto

@ExcelWritable
data class ExcelWriterValidationTypeTextLengthErrorDto(
  @ExcelWriterHeader(
    name = "SAMPLE TEXT LENGTH"
  )
  @ExcelWriterColumn(
    validationType = DataValidationConstraint.ValidationType.TEXT_LENGTH,
  )
  val text: String,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterValidationTypeTextLengthErrorDto> {
    override fun createSampleData(size: Int): Collection<ExcelWriterValidationTypeTextLengthErrorDto> {
      return (1..size).map {
        ExcelWriterValidationTypeTextLengthErrorDto(
          text = "text..."
        )
      }
    }
  }
}
