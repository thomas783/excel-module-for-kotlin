package reader.test

import org.excelkotlin.reader.ExcelReader
import org.excelkotlin.reader.ExcelReaderFieldError
import org.excelkotlin.reader.exception.ExcelReaderException
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import io.kotest.matchers.shouldBe
import io.kotest.matchers.types.shouldBeTypeOf
import reader.dto.ExcelReaderSampleDto
import shared.OrderStatus
import shared.getLocalPath
import java.time.LocalDate
import java.time.LocalDateTime

@OptIn(ExperimentalKotest::class)
class ExcelReaderCellTypeTests : BehaviorSpec({
  lateinit var excelReader: ExcelReader

  afterTest { excelReader.close() }

  given("Excel file for normal cases") {
    val localPath = getLocalPath("sample-to-read.xlsx")
    excelReader = ExcelReader(localPath)
    val excelData = excelReader.readExcelFile<ExcelReaderSampleDto>()

    given("class for columns are specified") {
      given("expected to be String(countryCode, sku, orderNumber, option, textLengthGreaterThanThree for this case)") {
        val expectedToBeString = excelData.map {
          with(it) { listOf(countryCode, sku, orderNumber, option, textLengthGreaterThanThree) }
        }.flatten().filterNotNull()

        then("should read cells for these columns as String") {
          expectedToBeString.forEach {
            debug { "cell value: $it, class: ${it.javaClass.simpleName}" }

            it.shouldBeTypeOf<String>()
          }
        }
      }

      given("expected to be Double(price, decimalBetween0And10 for this case)") {
        val expectedToBeDouble = excelData.map {
          with(it) { listOf(price, decimalBetween0And10) }
        }.flatten().filterNotNull()

        then("should read cells for these columns as Double") {
          expectedToBeDouble.forEach {
            debug { "cell value: $it, class: ${it.javaClass.simpleName}" }

            it.shouldBeTypeOf<Double>()
          }
        }
      }

      given("expected to be Int(quantity, integerGreaterThan5 for this case)") {
        val expectedToBeInt = excelData.map {
          with(it) { listOf(quantity, integerGreaterThan5) }
        }.flatten().filterNotNull()

        then("should read cells for these columns as Int") {
          expectedToBeInt.forEach {
            debug { "cell value: $it, class: ${it.javaClass.simpleName}" }

            it.shouldBeTypeOf<Int>()
          }
        }
      }

      given("expected to be LocalDateTime(orderedAt for this case)") {
        val expectedToBeLocalDateTime = excelData.mapNotNull { it.orderedAt }

        then("should read cells for this column as LocalDateTime") {
          expectedToBeLocalDateTime.forEach {
            debug { "cell value: $it, class: ${it.javaClass.simpleName}" }

            it.shouldBeTypeOf<LocalDateTime>()
          }
        }
      }

      given("expected to be LocalDate(paidDate for this case)") {
        val expectedToBeLocalDate = excelData.mapNotNull { it.paidDate }

        then("should read cells for this column as LocalDateTime") {
          expectedToBeLocalDate.forEach {
            debug { "cell value: $it, class: ${it.javaClass.simpleName}" }

            it.shouldBeTypeOf<LocalDate>()
          }
        }
      }

      given("expected to be specified enum class(OrderStatus for this case)") {
        val expectedToBeEnum = excelData.mapNotNull { it.orderStatus }

        then("should read cells for this column as OrderStatus") {
          expectedToBeEnum.forEach {
            debug { "cell value: $it, class: ${it.javaClass.simpleName}" }

            it.shouldBeTypeOf<OrderStatus>()
          }
        }
      }
    }
  }

  given("Excel file for wrong cell type cases") {
    val localPath = getLocalPath("sample-wrong-cell-type.xlsx")
    excelReader = ExcelReader(localPath)

    val exception = runCatching {
      excelReader.readExcelFile<ExcelReaderSampleDto>()
    }.exceptionOrNull() as ExcelReaderException

    given("expected to be String(countryCode, sku, orderNumber, option, textLengthGreaterThanThree for this case) are not String cell type") {
      then("exception for expected to be String columns should be ExcelReaderFieldError.TYPE ") {
        val stringInvalidTypeExceptions = exception.errorFieldList.filter {
          it.fieldHeader in listOf(
            "countryCode",
            "sku",
            "orderNumber",
            "option",
            "textLengthGreaterThanThree"
          )
        }

        debug { stringInvalidTypeExceptions.joinToString("\n") }

        stringInvalidTypeExceptions.forEach {
          it.type shouldBe ExcelReaderFieldError.TYPE.name
        }
      }
    }

    given("expected to be Numeric(price, decimalBetween0And10, quantity, integerGreaterThan5, orderedAt, paidDate for this case) are not Numeric cell type") {
      then("exception for expected to be numeric columns should be ExcelReaderFieldError.TYPE ") {
        val numericInvalidTypeExceptions = exception.errorFieldList.filter {
          it.fieldHeader in listOf(
            "price",
            "decimalBetween0And10",
            "quantity",
            "integerGreaterThan5",
            "orderedAt",
            "paidDate"
          )
        }

        debug { numericInvalidTypeExceptions.joinToString("\n") }

        numericInvalidTypeExceptions.forEach {
          it.type shouldBe ExcelReaderFieldError.TYPE.name
        }
      }
    }

    given("expected to be Enum(orderStatus) is not converted to enum class(OrderStatus for this case) expected") {
      then("exception for expected to be enum columns should be ExcelReaderFieldError.TYPE") {
        val enumInvalidTypeExceptions = exception.errorFieldList.filter { it.fieldHeader == "orderStatus" }

        debug { enumInvalidTypeExceptions.joinToString("\n") }

        enumInvalidTypeExceptions.forEach {
          it.type shouldBe ExcelReaderFieldError.TYPE.name
        }
      }
    }
  }
})
