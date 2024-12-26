package reader.dto

import com.excelkotlin.reader.IExcelReaderCommonDto
import com.excelkotlin.reader.annotation.ExcelReaderHeader
import org.valiktor.ConstraintViolationException
import org.valiktor.functions.hasSize
import org.valiktor.functions.isBetween
import org.valiktor.functions.isGreaterThanOrEqualTo
import org.valiktor.functions.isIn
import org.valiktor.functions.isNotNull
import org.valiktor.functions.matches
import org.valiktor.validate
import shared.OrderStatus
import java.time.LocalDate
import java.time.LocalDateTime

@ExcelReaderHeader(
  essentialFields = ["countryCode", "sku", "orderNumber", "orderStatus", "price", "quantity", "orderedAt"]
)
data class ExcelReaderSampleDto(
  var countryCode: String? = null,
  var sku: String? = null,
  var orderNumber: String? = null,
  var orderStatus: OrderStatus? = null,
  var price: Double? = null,
  var quantity: Int? = null,
  var orderedAt: LocalDateTime? = null,
  var paidDate: LocalDate? = null,
  var productName: String? = null,
  var option: String? = null,
  var textLengthGreaterThanThree: String? = null,
  var decimalBetween0And10: Double? = null,
  var integerGreaterThan5: Int? = null
) : IExcelReaderCommonDto {

  @Throws(ConstraintViolationException::class)
  override fun validate() {
    validate(this) {
      validate(ExcelReaderSampleDto::countryCode).isNotNull().matches(countryCodeRegex)
      validate(ExcelReaderSampleDto::sku).isNotNull()
      validate(ExcelReaderSampleDto::orderNumber).isNotNull()
      validate(ExcelReaderSampleDto::orderStatus).isNotNull()
      validate(ExcelReaderSampleDto::price).isNotNull().isGreaterThanOrEqualTo(0.0)
      validate(ExcelReaderSampleDto::quantity).isNotNull().isGreaterThanOrEqualTo(0)
      validate(ExcelReaderSampleDto::orderedAt).isNotNull()
      validate(ExcelReaderSampleDto::option).isIn("option1", "option2", "option3")
      validate(ExcelReaderSampleDto::textLengthGreaterThanThree).hasSize(min = 3)
      validate(ExcelReaderSampleDto::decimalBetween0And10).isBetween(0.0, 10.0)
      validate(ExcelReaderSampleDto::integerGreaterThan5).isGreaterThanOrEqualTo(5)
    }
  }

  companion object {
    val countryCodeRegex = "^[A-Z]{2}$".toRegex()
  }
}
