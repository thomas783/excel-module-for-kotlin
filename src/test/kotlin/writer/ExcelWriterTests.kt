package writer

import excel.writer.ExcelWriter
import excel.writer.annotation.ExcelWriterColumn
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import writer.dto.ExcelWriterSampleDto
import java.io.File
import java.io.FileOutputStream
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties

internal class ExcelWriterTests : BehaviorSpec({
  val sampleDataSize = 1000
  val sheetName = "Sample Excel File"
  val localPath = "src/test/resources/sample.xlsx"
  val out = FileOutputStream(localPath)

  lateinit var createdFile: File
  lateinit var createdWorkbook: Workbook

  beforeSpec {
    val sampleData = ExcelWriterSampleDto.createSampleData(sampleDataSize)
    val workbook = ExcelWriter.createWorkbook(sampleData, sheetName)
    workbook.write(out)
    out.close()
    workbook.dispose()
    workbook.close()
    createdFile = File(localPath)
  }

  beforeTest {
    createdWorkbook = createdFile.inputStream().use { WorkbookFactory.create(it) }
  }

  afterTest {
    createdWorkbook.close()
  }

  given("ExcelWriter") {
    `when`("create workbook") {
      then("excel file well created") {
        createdFile.exists() shouldBe true
      }
      val sheet = createdWorkbook.getSheetAt(0)
      then("workbook sheet name well created as annotated") {
        sheet.sheetName shouldBe sheetName
      }
      then("workbook sheet has ${sampleDataSize + 1} rows including header row") {
        sheet.physicalNumberOfRows shouldBe sampleDataSize + 1
      }
      val headerRow = sheet.getRow(0)
      then("header row cell counts equal to ExcelWriterSampleDto properties counts") {
        val excelWriterSampleDtoPropertiesCounts = ExcelWriterSampleDto::class.memberProperties.size
        headerRow.physicalNumberOfCells shouldBe excelWriterSampleDtoPropertiesCounts
      }
      then("header row cell values well created as constructors in order") {
        val memberPropertiesMap = ExcelWriterSampleDto::class.memberProperties.associate {
          it.name to it.findAnnotation<ExcelWriterColumn>()
        }
        val excelWriterSampleDtoConstructorNamesInOrder = ExcelWriterSampleDto::class.constructors.map { constructor ->
          constructor.parameters.map { parameter ->
            memberPropertiesMap[parameter.name]?.headerName ?: parameter.name
          }
        }.flatten().also {
          println("ExcelWriterSampleDto constructor names in order: $it")
        }
        val headerRowCellValues = (0 until headerRow.physicalNumberOfCells).map {
          headerRow.getCell(it).stringCellValue
        }.also { println("Excel file header row cell values: $it") }
        (0 until headerRow.physicalNumberOfCells).forEach {
          headerRowCellValues[it] shouldBe excelWriterSampleDtoConstructorNamesInOrder[it]
        }
      }
    }
  }

  afterSpec {
    createdFile.delete()
  }
})
