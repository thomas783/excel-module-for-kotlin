package reader.test

import com.excelkotlin.reader.ExcelReader
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import reader.dto.ExcelReaderSampleDto

@OptIn(ExperimentalKotest::class)
class ExcelReaderTests : BehaviorSpec({
  val localPath = "src/test/resources/sample/sample-to-read.xlsx"

  given("Base excel file") {
    val excelData = ExcelReader(localPath).readExcelFile<ExcelReaderSampleDto>()

    then("Excel file is read successfully") {
      debug { "Excel data: $excelData" }
    }
  }
})
