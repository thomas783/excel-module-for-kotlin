package writer.dto.validationtypeerror

import excel.writer.annotation.ExcelWriterColumn
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.IndexedColors
import shared.IExcelWriterCommonDto
import shared.OrderStatus

data class ExcelWriterValidationTypeListErrorDto(
  @ExcelWriterColumn(
    headerName = "ORDER STATUS",
    headerCellColor = IndexedColors.RED,
    validationType = DataValidationConstraint.ValidationType.LIST
  )
  val orderStatus: OrderStatus,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterValidationTypeListErrorDto> {
    override fun createSampleData(size: Int): List<ExcelWriterValidationTypeListErrorDto> {
      return (1..size).map {
        ExcelWriterValidationTypeListErrorDto(
          orderStatus = OrderStatus.entries.toTypedArray().random()
        )
      }
    }
  }
}
