package writer.dto.validation

import com.excelkotlin.writer.annotation.ExcelWritable
import com.excelkotlin.writer.annotation.ExcelWriterColumn
import com.excelkotlin.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import writer.dto.IExcelWriterCommonDto

@ExcelWritable
data class ExcelWriterValidationTypeFormulaErrorDto(
  @ExcelWriterHeader(name = "SAMPLE FORMULA")
  @ExcelWriterColumn(
    validationType = DataValidationConstraint.ValidationType.FORMULA,
  )
  val formula: String,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterValidationTypeFormulaErrorDto> {
    override fun createSampleData(size: Int): List<ExcelWriterValidationTypeFormulaErrorDto> {
      return (1..size).map {
        ExcelWriterValidationTypeFormulaErrorDto(
          formula = "formula expected..."
        )
      }
    }
  }
}
