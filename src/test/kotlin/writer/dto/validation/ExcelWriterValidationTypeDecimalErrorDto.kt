package writer.dto.validation

import com.excelkotlin.writer.annotation.ExcelWritable
import com.excelkotlin.writer.annotation.ExcelWriterColumn
import com.excelkotlin.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

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
