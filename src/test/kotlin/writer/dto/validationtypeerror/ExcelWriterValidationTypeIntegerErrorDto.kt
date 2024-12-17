package writer.dto.validationtypeerror

import excel.writer.annotation.ExcelWriterColumn
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

data class ExcelWriterValidationTypeIntegerErrorDto(
  @ExcelWriterColumn(
    headerName = "SAMPLE INTEGER",
    validationType = DataValidationConstraint.ValidationType.INTEGER
  )
  val integer: Int,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterValidationTypeIntegerErrorDto> {
    override fun createSampleData(size: Int): Collection<ExcelWriterValidationTypeIntegerErrorDto> {
      return (1..size).map {
        ExcelWriterValidationTypeIntegerErrorDto(
          integer = it
        )
      }
    }
  }
}
