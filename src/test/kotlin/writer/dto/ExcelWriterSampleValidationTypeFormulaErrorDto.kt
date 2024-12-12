package writer.dto

import excel.writer.annotation.ExcelWriterColumn
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.IExcelWriterCommonDto

data class ExcelWriterSampleValidationTypeFormulaErrorDto(
  @ExcelWriterColumn(
    headerName = "SAMPLE FORMULA",
    validationType = DataValidationConstraint.ValidationType.FORMULA,
  )
  val formula: String,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterSampleValidationTypeFormulaErrorDto> {
    override fun createSampleData(size: Int): List<ExcelWriterSampleValidationTypeFormulaErrorDto> {
      return (1..size).map {
        ExcelWriterSampleValidationTypeFormulaErrorDto(
          formula = "formula expected..."
        )
      }
    }
  }
}
