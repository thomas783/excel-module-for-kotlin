package reader.test

import com.excelkotlin.reader.ExcelReader
import com.excelkotlin.reader.annotation.ExcelReaderHeader
import com.excelkotlin.reader.exception.ExcelReaderFileExtensionException
import io.kotest.assertions.throwables.shouldThrow
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import io.kotest.matchers.collections.shouldBeIn
import io.kotest.matchers.shouldBe
import reader.dto.ExcelReaderSampleDto
import kotlin.reflect.full.findAnnotation

@OptIn(ExperimentalKotest::class)
class ExcelReaderBaseTests : BehaviorSpec({
  lateinit var excelReader: ExcelReader

  afterTest {
    excelReader.close()
  }

  given("Excel file for basic tests") {
    val localPath = getLocalPath("sample-to-read.xlsx")
    excelReader = ExcelReader(localPath)
    val excelData = excelReader.readExcelFile<ExcelReaderSampleDto>()

    then("Excel file is read successfully") {
      debug { "Excel data: $excelData" }

      excelData.size shouldBe 1000
    }

    `when`("ExcelReaderHeader annotation essential fields are not empty") {
      then("file header should be containing all columns in essential fields") {
        val headerRow = excelReader.getHeader<ExcelReaderSampleDto>().map { it.key }
        val essentialFields = ExcelReaderSampleDto::class.findAnnotation<ExcelReaderHeader>()?.essentialFields

        debug { "headerRow: $headerRow" }

        essentialFields?.forEach { essentialFieldName ->
          debug { "essentialFieldName: $essentialFieldName" }

          essentialFieldName shouldBeIn headerRow
        }
      }
    }
  }

  given("Excel file for wrong extension") {
    val localPath = getLocalPath("sample-to-read.csv")

    then("ExcelReaderFileExtensionException is thrown") {
      shouldThrow<ExcelReaderFileExtensionException> {
        debug { "path: $localPath" }

        ExcelReader(localPath)
      }
    }
  }
})

fun getLocalPath(path: String): String {
  return "src/test/resources/sample/$path"
}
