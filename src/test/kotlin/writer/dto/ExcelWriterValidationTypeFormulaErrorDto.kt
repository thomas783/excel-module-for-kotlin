package writer.dto

import excel.writer.annotation.ExcelWriterColumn
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

data class ExcelWriterValidationTypeFormulaErrorDto(
  @ExcelWriterColumn(
    headerName = "SAMPLE FORMULA",
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
