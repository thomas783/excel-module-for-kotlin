package writer.test

import com.excelkotlin.writer.ExcelWriter
import com.excelkotlin.writer.exception.ExcelWritableMissingException
import io.kotest.assertions.throwables.shouldThrow
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import io.kotest.matchers.shouldBe
import shared.ExcelWriterBaseTests.Companion.setExcelWriterCommonSpec
import writer.dto.ExcelWriterSampleDto
import writer.dto.ExcelWriterWritableMissingErrorDto

@OptIn(ExperimentalKotest::class)
internal class ExcelWriterCreateTests : BehaviorSpec({
  val sampleDataSize = 1000
  val baseTest = setExcelWriterCommonSpec<ExcelWriterSampleDto.Companion, ExcelWriterSampleDto>(
    sampleDataSize = sampleDataSize,
    path = "sample-create",
  )

  given("ExcelWritable Annotation") {
    then("excel file is created") {
      baseTest.excelFile.exists() shouldBe true
    }

    val sheet = baseTest.workbook.getSheetAt(0)

    then("workbook sheet name well created as expected") {

      debug { "created sheet name: ${sheet.sheetName}" }
      debug { "expected sheet name: ${baseTest.sheetName}" }

      sheet.sheetName shouldBe baseTest.sheetName
    }

    then("workbook sheet has ${sampleDataSize + 1} rows including header row") {

      debug { "created sheet row count: ${sheet.physicalNumberOfRows}" }
      debug { "expected sheet row count: ${sampleDataSize + 1}" }

      sheet.physicalNumberOfRows shouldBe sampleDataSize + 1
    }
  }

  given("ExcelWritable Annotation is missing") {
    val excelWritableMissingErrorDto = ExcelWriterWritableMissingErrorDto.createSampleData(sampleDataSize)

    then("ExcelWritableMissingException is thrown") {
      shouldThrow<ExcelWritableMissingException> {
        ExcelWriter.createWorkbook(excelWritableMissingErrorDto, "sample-create-missing-annotation")
      }.also { debug { it } }
    }
  }
})
