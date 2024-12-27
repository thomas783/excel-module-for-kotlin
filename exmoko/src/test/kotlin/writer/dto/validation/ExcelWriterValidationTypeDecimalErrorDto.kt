package writer.dto.validation

import org.exmoko.writer.annotation.ExcelWritable
import org.exmoko.writer.annotation.ExcelWriterColumn
import org.exmoko.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import writer.dto.IExcelWriterCommonDto

@ExcelWritable
data class ExcelWriterValidationTypeDecimalErrorDto(
  @ExcelWriterHeader(
    name = "SAMPLE DECIMAL"
  )
  @ExcelWriterColumn(
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
