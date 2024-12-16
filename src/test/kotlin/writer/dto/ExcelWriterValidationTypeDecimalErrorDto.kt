package writer.dto

import excel.writer.annotation.ExcelWriterColumn
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

data class ExcelWriterValidationTypeDecimalErrorDto(
  @ExcelWriterColumn(
    headerName = "SAMPLE DECIMAL",
    validationType = DataValidationConstraint.ValidationType.DECIMAL
  )
  val decimal: Double,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterValidationTypeDecimalErrorDto> {
    override fun createSampleData(size: Int): Collection<ExcelWriterValidationTypeDecimalErrorDto> {
      return (1..size).map {
        ExcelWriterValidationTypeDecimalErrorDto(
          decimal = 1.0 * it
        )
      }
    }
  }
}
