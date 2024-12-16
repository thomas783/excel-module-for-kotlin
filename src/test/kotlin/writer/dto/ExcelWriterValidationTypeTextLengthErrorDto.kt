package writer.dto

import excel.writer.annotation.ExcelWriterColumn
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

data class ExcelWriterValidationTypeTextLengthErrorDto(
  @ExcelWriterColumn(
    headerName = "SAMPLE TEXT LENGTH",
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
