package specs

import dto.ExcelWriterSampleDto
import excel.writer.ExcelWriter
import io.kotest.core.spec.style.ShouldSpec
import org.apache.poi.xssf.streaming.SXSSFWorkbook

internal class ExcelWriterShouldSpec : ShouldSpec({
  val sampleDataSize = 1000
  val sheetName = "Sample Excel File"

  lateinit var sampleData: List<ExcelWriterSampleDto>
  lateinit var workbook: SXSSFWorkbook

  beforeSpec {
    sampleData = ExcelWriterSampleDto.createSampleData(sampleDataSize)
    workbook = ExcelWriter.createWorkbook(sampleData, sheetName)

  }

  should("workbook sheet well created") {
    assert(workbook.getSheetAt(0).sheetName == sheetName)
  }

  afterSpec { workbook.close() }
})
