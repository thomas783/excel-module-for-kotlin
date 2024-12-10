package writer

import excel.writer.annotation.ExcelWriterColumn
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.info
import io.kotest.matchers.shouldBe
import shared.ExcelWriterBaseTests
import writer.dto.ExcelWriterSampleDto
import kotlin.reflect.full.findAnnotation

@OptIn(ExperimentalKotest::class)
internal class ExcelWriterHeaderRowTests : BehaviorSpec({
  val sampleDataSize = 1000
  val sampleData = ExcelWriterSampleDto.createSampleData(sampleDataSize)
  val baseTest = ExcelWriterBaseTests().also {
    it.setCommonSpec(this, sampleData)
  }

  given("ExcelWriterColumn Annotation") {
    val sheet = baseTest.createdWorkbook.getSheetAt(0)
    val headerRow = sheet.getRow(0)
    `when`("annotation is provided in constructor") {

      then("header row cell counts equal to ExcelWriterSampleDto properties counts that has ExcelWriterColumn annotation") {
        val excelWriterSampleDtoPropertiesCounts = ExcelWriterSampleDto.getMemberProperties().size
        headerRow.physicalNumberOfCells shouldBe excelWriterSampleDtoPropertiesCounts
      }

      then("header row cell values well created as constructors in order") {
        val memberPropertiesMap = ExcelWriterSampleDto.getMemberProperties().associate {
          it.name to it.findAnnotation<ExcelWriterColumn>()
        }
        val excelWriterSampleDtoConstructorNamesInOrder =
          ExcelWriterSampleDto.getConstructorParameters().mapNotNull { parameter ->
            val excelWriterColumn = memberPropertiesMap[parameter.name]
            val headerName = excelWriterColumn?.headerName
            headerName?.ifBlank { parameter.name }
          }

        info { "ExcelWriterSampleDto constructor names in order: $excelWriterSampleDtoConstructorNamesInOrder" }

        val headerRowCellValues = (0 until headerRow.physicalNumberOfCells).map {
          headerRow.getCell(it).stringCellValue
        }

        info { "Excel file header row cell values: $headerRowCellValues" }

        (0 until headerRow.physicalNumberOfCells).forEach {
          headerRowCellValues[it] shouldBe excelWriterSampleDtoConstructorNamesInOrder[it]
        }
      }
    }
  }
})
