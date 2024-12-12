package writer.dto

import excel.writer.annotation.ExcelWriterColumn
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.IndexedColors
import shared.IExcelWriterCommonDto
import shared.OrderStatus

data class ExcelWriterSampleValidationTypeListErrorDto(
  @ExcelWriterColumn(
    headerName = "ORDER STATUS",
    headerCellColor = IndexedColors.RED,
    validationType = DataValidationConstraint.ValidationType.LIST
  )
  val orderStatus: OrderStatus,
) {
  companion object : IExcelWriterCommonDto<ExcelWriterSampleValidationTypeListErrorDto> {
    override fun createSampleData(size: Int): List<ExcelWriterSampleValidationTypeListErrorDto> {
      return (1..size).map {
        ExcelWriterSampleValidationTypeListErrorDto(
          orderStatus = OrderStatus.entries.toTypedArray().random()
        )
      }
    }
  }
}
