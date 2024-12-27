package writer.dto.validation

import org.exmoko.writer.annotation.ExcelWritable
import org.exmoko.writer.annotation.ExcelWriterColumn
import org.exmoko.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import writer.dto.IExcelWriterCommonDto

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
