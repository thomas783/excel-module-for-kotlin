package writer

import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.ShouldSpec
import io.kotest.engine.test.logging.info
import io.kotest.matchers.shouldBe
import shared.ExcelWriterBaseTests.Companion.setExcelWriterCommonSpec
import writer.dto.ExcelWriterSampleDto

@OptIn(ExperimentalKotest::class)
internal class ExcelWriterCreateTests : ShouldSpec({
  val sampleDataSize = 1000
  val baseTest = setExcelWriterCommonSpec<ExcelWriterSampleDto.Companion, ExcelWriterSampleDto>(
    sampleDataSize = sampleDataSize,
    path = "sample-create",
  )

  should("workbook sheet well created") {
    baseTest.excelFile.exists() shouldBe true
  }

  should("workbook sheet name well created as annotated") {
    val sheet = baseTest.workbook.getSheetAt(0)

    info { "created sheet name: ${sheet.sheetName}" }
    info { "expected sheet name: ${baseTest.sheetName}" }

    sheet.sheetName shouldBe baseTest.sheetName
  }

  should("workbook sheet has ${sampleDataSize + 1} rows including header row") {
    val sheet = baseTest.workbook.getSheetAt(0)

    info { "created sheet row count: ${sheet.physicalNumberOfRows}" }
    info { "expected sheet row count: ${sampleDataSize + 1}" }

    sheet.physicalNumberOfRows shouldBe sampleDataSize + 1
  }
})
