package reader.test

import org.exmoko.reader.ExcelReader
import org.exmoko.reader.ExcelReaderFieldError
import org.exmoko.reader.exception.ExcelReaderException
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import io.kotest.matchers.shouldBe
import reader.dto.ExcelReaderSampleDto
import shared.getLocalPath

@OptIn(ExperimentalKotest::class)
class ExcelReaderCellValueTests : BehaviorSpec({
  lateinit var excelReader: ExcelReader

  afterTest { excelReader.close() }

  given("Excel file for wrong validation cases") {
    val localPath = getLocalPath("sample-wrong-validation.xlsx")
    excelReader = ExcelReader(localPath)

    val exception = runCatching {
      excelReader.readExcelFile<ExcelReaderSampleDto>()
    }.exceptionOrNull() as ExcelReaderException


    given("validation for countryCode is not null and matches regex") {
      val countryCodeExceptions = exception.errorFieldList.filter { it.fieldHeader == "countryCode" }

      then("exception type for countryCode should be ExcelReaderFieldError.VALID") {
        countryCodeExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for sku is not null") {
      val skuExceptions = exception.errorFieldList.filter { it.fieldHeader == "sku" }

      then("exception type for sku should be ExcelReaderFieldError.VALID") {
        skuExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for orderNumber is not null") {
      val orderNumberExceptions = exception.errorFieldList.filter { it.fieldHeader == "orderNumber" }

      then("exception type for orderNumber should be ExcelReaderFieldError.VALID") {
        orderNumberExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for orderStatus is not null") {
      val orderStatusExceptions = exception.errorFieldList.filter { it.fieldHeader == "orderStatus" }

      then("exception type for orderStatus should be ExcelReaderFieldError.VALID") {
        orderStatusExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for price is not null and greater than or equal to 0.0") {
      val priceExceptions = exception.errorFieldList.filter { it.fieldHeader == "price" }

      then("exception type for orderStatus should be ExcelReaderFieldError.VALID") {
        priceExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for quantity is not null and greater than or equal to 0") {
      val quantityExceptions = exception.errorFieldList.filter { it.fieldHeader == "quantity" }

      then("exception type for quantity should be ExcelReaderFieldError.VALID") {
        quantityExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for orderedAt is not null") {
      val orderedAtExceptions = exception.errorFieldList.filter { it.fieldHeader == "orderedAt" }

      then("exception type for orderedAt should be ExcelReaderFieldError.VALID") {
        orderedAtExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for option is in option1, option2, option3") {
      val optionExceptions = exception.errorFieldList.filter { it.fieldHeader == "option" }

      then("exception type for option should be ExcelReaderFieldError.VALID") {
        optionExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for textLengthGreaterThanThree has size greater than or equal to 3") {
      val textLengthGreaterThanThreeExceptions = exception.errorFieldList.filter { it.fieldHeader == "textLengthGreaterThanThree" }

      then("exception type for textLengthGreaterThanThree should be ExcelReaderFieldError.VALID") {
        textLengthGreaterThanThreeExceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for decimalBetween0And10 is between 0.0 and 10.0") {
      val decimalBetween0And10Exceptions = exception.errorFieldList.filter { it.fieldHeader == "decimalBetween0And10" }

      then("exception type for decimalBetween0And10 should be ExcelReaderFieldError.VALID") {
        decimalBetween0And10Exceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }

    given("validation for integerGreaterThan5 is greater than or equal to 5") {
      val integerGreaterThan5Exceptions = exception.errorFieldList.filter { it.fieldHeader == "integerGreaterThan5" }

      then("exception type for integerGreaterThan5 should be ExcelReaderFieldError.VALID") {
        integerGreaterThan5Exceptions.forEach {
          debug { it }
          it.type shouldBe ExcelReaderFieldError.VALID.name
        }
      }
    }
  }
})
