package writer

import io.kotest.core.spec.style.ShouldSpec
import io.kotest.matchers.shouldBe
import shared.ExcelWriterBaseTests
import writer.dto.ExcelWriterSampleDto

internal class ExcelWriterCreateTests : ShouldSpec({
  val sampleDataSize = 1000
  val sampleData = ExcelWriterSampleDto.createSampleData(sampleDataSize)
  val baseTest = ExcelWriterBaseTests().also {
    it.setCommonSpec(this, sampleData)
  }

  should("workbook sheet well created") {
    baseTest.createdFile.exists() shouldBe true
  }


  should("workbook sheet name well created as annotated") {
    val sheet = baseTest.createdWorkbook.getSheetAt(0)
    sheet.sheetName shouldBe baseTest.sheetName
  }

  should("workbook sheet has ${sampleDataSize + 1} rows including header row") {
    val sheet = baseTest.createdWorkbook.getSheetAt(0)
    sheet.physicalNumberOfRows shouldBe sampleDataSize + 1
  }
})
