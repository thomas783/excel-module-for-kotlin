package writer

import excel.writer.ExcelWriter
import io.kotest.core.spec.style.ShouldSpec
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import writer.dto.ExcelWriterSampleDto

internal class ExcelTestExample : ShouldSpec({
  val sampleDataSize = 1000
  val sheetName = "Sample Excel File"
  val sampleData = ExcelWriterSampleDto.createSampleData(sampleDataSize)
  val workbook = ExcelWriter.createWorkbook(sampleData, sheetName)

  beforeSpec {
    val readableWorkbook = WorkbookFactory.create(workbook)
  }

})
