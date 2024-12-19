package writer.dto.validation

import excel.writer.annotation.ExcelWritable
import excel.writer.annotation.ExcelWriterColumn
import excel.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

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
