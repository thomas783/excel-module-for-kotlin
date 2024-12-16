package writer.dto

import excel.writer.annotation.ExcelWriterColumn
import org.apache.poi.ss.usermodel.DataValidation
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.IndexedColors
import shared.IExcelWriterCommonDto
import shared.OrderStatus
import java.time.LocalDate
import java.time.LocalDateTime

data class ExcelWriterSampleDto(
  @ExcelWriterColumn(
    headerName = "COUNTRY CODE",
    headerCellColor = IndexedColors.RED,
    validationType = DataValidationConstraint.ValidationType.FORMULA,
    validationIgnoreBlank = false,
    validationFormula = "AND(EXACT(UPPER(${ExcelWriterColumn.CURRENT_CELL}), ${ExcelWriterColumn.CURRENT_CELL}), LEN(${ExcelWriterColumn.CURRENT_CELL}) = 2)",
    validationPromptTitle = "COUNTRY CODE",
    validationErrorStyle = DataValidation.ErrorStyle.STOP,
    validationErrorTitle = "Invalid country code format",
    validationErrorText = "Country code should be two uppercase alphabets. Example: KR,JP,US...",
  )
  val countryCode: String,

  @ExcelWriterColumn(
    headerCellColor = IndexedColors.RED,
  )
  val sku: String,

  @ExcelWriterColumn(
    headerName = "ORDER NUMBER",
    headerCellColor = IndexedColors.RED,
    validationPromptTitle = "ORDER NUMBER"
  )
  val orderNumber: String,

  @ExcelWriterColumn(
    headerName = "ORDER STATUS",
    headerCellColor = IndexedColors.RED,
    validationType = DataValidationConstraint.ValidationType.LIST,
    validationIgnoreBlank = false,
    validationErrorStyle = DataValidation.ErrorStyle.STOP,
    validationListEnum = OrderStatus::class,
    validationPromptTitle = "ORDER STATUS",
    validationErrorTitle = "Invalid order status format",
  )
  val orderStatus: OrderStatus,

  @ExcelWriterColumn(
    headerName = "PRICE",
    headerCellColor = IndexedColors.RED,
    validationPromptTitle = "PRICE"
  )
  val price: Double,

  @ExcelWriterColumn(
    headerName = "QUANTITY",
    headerCellColor = IndexedColors.RED,
    validationPromptTitle = "QUANTITY"
  )
  val quantity: Int,

  @ExcelWriterColumn(
    headerName = "ORDERED AT",
    headerCellColor = IndexedColors.BLUE,
    validationPromptTitle = "ORDERED AT"
  )
  val orderedAt: LocalDateTime? = null,

  @ExcelWriterColumn(
    headerName = "PAID DATE",
    headerCellColor = IndexedColors.BLUE,
    validationPromptTitle = "PAID DATE"
  )
  val paidDate: LocalDate? = null,

  @ExcelWriterColumn(
    headerName = "PRODUCT NAME",
    validationPromptTitle = "PRODUCT NAME"
  )
  val productName: String? = null,

  @ExcelWriterColumn(
    headerName = "SAMPLE LIST",
    validationType = DataValidationConstraint.ValidationType.LIST,
    validationListOptions = ["option1", "option2", "option3"]
  )
  val option: String,

  @ExcelWriterColumn(
    headerName = "TEXT LENGTH LIMIT TO THREE",
    validationType = DataValidationConstraint.ValidationType.TEXT_LENGTH,
    operationType = DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
    operationFormula1 = "3",
  )
  val textLengthGreaterThanThree: String? = null,

  @ExcelWriterColumn(
    headerName = "DECIMAL BETWEEN 0 AND 10",
    validationType = DataValidationConstraint.ValidationType.DECIMAL,
    operationType = DataValidationConstraint.OperatorType.BETWEEN,
    operationFormula1 = "0",
    operationFormula2 = "10"
  )
  val decimalBetween0And10: Double,

  @ExcelWriterColumn(
    headerName = "INTEGER GREATER THAN 5",
    validationType = DataValidationConstraint.ValidationType.INTEGER,
    operationType = DataValidationConstraint.OperatorType.GREATER_THAN,
    operationFormula1 = "5"
  )
  val integerGreaterThan5: Int,

  val extraField: String? = null,
) {

  companion object : IExcelWriterCommonDto<ExcelWriterSampleDto> {
    override fun createSampleData(size: Int): List<ExcelWriterSampleDto> {
      return (1..size).map { number ->
        ExcelWriterSampleDto(
          countryCode = "KR",
          sku = "SKU-$number",
          orderNumber = "orderNumber-$number",
          orderStatus = OrderStatus.entries.toTypedArray().random(),
          price = (number % 10) * 1000.0,
          quantity = number % 3 + 1,
          orderedAt = LocalDateTime.now().minusSeconds(number.toLong()),
          paidDate = LocalDate.now().minusDays((number % 3).toLong()),
          productName = "Product $number",
          option = "option${(1..3).random()}",
          decimalBetween0And10 = (number % 10).toDouble(),
          integerGreaterThan5 = number % 10 + 5,
        )
      }
    }
  }
}
