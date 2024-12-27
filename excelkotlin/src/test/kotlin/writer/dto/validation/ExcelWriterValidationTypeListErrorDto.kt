package writer.dto.validation

import org.excelkotlin.writer.annotation.ExcelWritable
import org.excelkotlin.writer.annotation.ExcelWriterColumn
import org.excelkotlin.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.IndexedColors
import writer.dto.IExcelWriterCommonDto
import shared.OrderStatus

@ExcelWritable
data class ExcelWriterValidationTypeListErrorDto(
  @ExcelWriterHeader(
    name = "ORDER STATUS",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(
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
