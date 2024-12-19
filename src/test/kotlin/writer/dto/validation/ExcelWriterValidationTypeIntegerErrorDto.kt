package writer.dto.validation

import excel.writer.annotation.ExcelWritable
import excel.writer.annotation.ExcelWriterColumn
import excel.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

@ExcelWritable
data class ExcelWriterValidationTypeIntegerErrorDto(
  @ExcelWriterHeader(
    name = "SAMPLE INTEGER"
  )
  @ExcelWriterColumn(
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
