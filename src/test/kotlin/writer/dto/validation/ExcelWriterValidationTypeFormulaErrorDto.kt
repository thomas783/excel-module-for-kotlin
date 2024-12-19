package writer.dto.validation

import excel.writer.annotation.ExcelWritable
import excel.writer.annotation.ExcelWriterColumn
import excel.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

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
