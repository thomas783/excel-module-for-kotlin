package writer

import io.kotest.core.spec.style.ShouldSpec
import io.kotest.matchers.shouldBe
import shared.ExcelWriterBaseTests.Companion.setCommonSpec
import writer.dto.ExcelWriterSampleDto

internal class ExcelWriterCreateTests : ShouldSpec({
  val sampleDataSize = 1000
  val baseTest = setCommonSpec<ExcelWriterSampleDto.Companion, ExcelWriterSampleDto>(
    sampleDataSize = sampleDataSize,
    path = "sample-create",
  )

  should("workbook sheet well created") {
    baseTest.excelFile.exists() shouldBe true
  }

  should("workbook sheet name well created as annotated") {
    val sheet = baseTest.workbook.getSheetAt(0)
    sheet.sheetName shouldBe baseTest.sheetName
  }

  should("workbook sheet has ${sampleDataSize + 1} rows including header row") {
    val sheet = baseTest.workbook.getSheetAt(0)
    sheet.physicalNumberOfRows shouldBe sampleDataSize + 1
  }
})
