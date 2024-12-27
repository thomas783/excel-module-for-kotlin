package writer.dto

import org.exmoko.writer.annotation.ExcelWritable
import org.exmoko.writer.annotation.ExcelWriterColumn
import org.exmoko.writer.annotation.ExcelWriterFreezePane
import org.exmoko.writer.annotation.ExcelWriterHeader
import org.apache.poi.ss.usermodel.DataValidation
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.IndexedColors
import shared.OrderStatus
import java.time.LocalDate
import java.time.LocalDateTime

@ExcelWritable(
  properties = [
    "countryCode", "sku", "orderNumber", "orderStatus", "price", "quantity",
    "orderedAt", "paidDate", "productName", "option", "textLengthGreaterThanThree",
    "decimalBetween0And10", "integerGreaterThan5"
  ]
)
@ExcelWriterFreezePane(rowSplit = 1)
data class ExcelWriterSampleDto(
  @ExcelWriterHeader(
    name = "COUNTRY CODE",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(
    validationType = DataValidationConstraint.ValidationType.FORMULA,
    validationIgnoreBlank = false,
    validationFormula = "AND(EXACT(UPPER(${ExcelWriterColumn.CURRENT_CELL}), ${ExcelWriterColumn.CURRENT_CELL}), LEN(${ExcelWriterColumn.CURRENT_CELL}) = 2)",
    validationPromptTitle = "COUNTRY CODE",
    validationErrorStyle = DataValidation.ErrorStyle.STOP,
    validationErrorTitle = "Invalid country code format",
    validationErrorText = "Country code should be two uppercase alphabets. Example: KR,JP,US...",
  )
  val countryCode: String,

  @ExcelWriterHeader(
    cellColor = IndexedColors.RED
  )
  val sku: String,

  @ExcelWriterHeader(
    name = "ORDER NUMBER",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(validationPromptTitle = "ORDER NUMBER")
  val orderNumber: String,

  @ExcelWriterHeader(
    name = "ORDER STATUS",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(
    validationType = DataValidationConstraint.ValidationType.LIST,
    validationIgnoreBlank = false,
    validationErrorStyle = DataValidation.ErrorStyle.STOP,
    validationListEnum = OrderStatus::class,
    validationPromptTitle = "ORDER STATUS",
    validationErrorTitle = "Invalid order status format",
  )
  val orderStatus: OrderStatus,

  @ExcelWriterHeader(
    name = "PRICE",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(
    validationPromptTitle = "PRICE"
  )
  val price: Double,

  @ExcelWriterHeader(
    name = "QUANTITY",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(
    validationPromptTitle = "QUANTITY"
  )
  val quantity: Int,

  @ExcelWriterHeader(
    name = "ORDERED AT",
    cellColor = IndexedColors.BLUE
  )
  @ExcelWriterColumn(
    validationPromptTitle = "ORDERED AT"
  )
  val orderedAt: LocalDateTime? = null,

  @ExcelWriterHeader(
    name = "PAID DATE",
    cellColor = IndexedColors.BLUE
  )
  @ExcelWriterColumn(
    validationPromptTitle = "PAID DATE"
  )
  val paidDate: LocalDate? = null,

  @ExcelWriterColumn(
    validationPromptTitle = "PRODUCT NAME"
  )
  val productName: String? = null,

  @ExcelWriterHeader(
    name = "SAMPLE LIST",
  )
  @ExcelWriterColumn(
    validationType = DataValidationConstraint.ValidationType.LIST,
    validationListOptions = ["option1", "option2", "option3"]
  )
  val option: String,

  @ExcelWriterHeader(
    name = "TEXT LENGTH LIMIT TO THREE",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(
    validationType = DataValidationConstraint.ValidationType.TEXT_LENGTH,
    operationType = DataValidationConstraint.OperatorType.GREATER_OR_EQUAL,
    operationFormula1 = "3",
  )
  val textLengthGreaterThanThree: String? = null,

  @ExcelWriterHeader(
    name = "DECIMAL BETWEEN 0 AND 10",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(
    validationType = DataValidationConstraint.ValidationType.DECIMAL,
    operationType = DataValidationConstraint.OperatorType.BETWEEN,
    operationFormula1 = "0",
    operationFormula2 = "10"
  )
  val decimalBetween0And10: Double,

  @ExcelWriterHeader(
    name = "INTEGER GREATER THAN 5",
    cellColor = IndexedColors.RED
  )
  @ExcelWriterColumn(
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
          orderStatus = OrderStatus.entries[number % OrderStatus.entries.size],
          price = (number % 10) * 1000.0,
          quantity = number % 3 + 1,
          orderedAt = LocalDate.now().atStartOfDay().plusSeconds(number.toLong()),
          paidDate = LocalDate.now().minusDays((number % 3).toLong()),
          productName = "Product $number",
          option = "option${number % 3 + 1}",
          decimalBetween0And10 = (number % 10).toDouble(),
          integerGreaterThan5 = number % 10 + 5,
        )
      }
    }
  }
}
