package writer

import excel.writer.annotation.ExcelWriterColumn
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.info
import io.kotest.matchers.shouldBe
import shared.ExcelWriterBaseTests.Companion.setExcelWriterCommonSpec
import writer.dto.ExcelWriterSampleDto
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.hasAnnotation
import kotlin.reflect.full.memberProperties

@OptIn(ExperimentalKotest::class)
internal class ExcelWriterHeaderRowTests : BehaviorSpec({
  val sampleDataKClass = ExcelWriterSampleDto::class
  val baseTest = setExcelWriterCommonSpec<ExcelWriterSampleDto.Companion, ExcelWriterSampleDto>(
    sampleDataSize = 1000,
    path = "sample-header-row",
  )

  given("ExcelWriterColumn Annotation") {
    val sheet = baseTest.workbook.getSheetAt(0)
    val headerRow = sheet.getRow(0)
    `when`("annotation is provided in constructor") {

      then("header row cell counts equal to ExcelWriterSampleDto properties counts that has ExcelWriterColumn annotation") {
        val excelWriterSampleDtoPropertiesCounts = sampleDataKClass.memberProperties.filter {
          it.hasAnnotation<ExcelWriterColumn>()
        }.size
        headerRow.physicalNumberOfCells shouldBe excelWriterSampleDtoPropertiesCounts
      }

      then("header row cell values well created as constructors in order") {
        val memberPropertiesMap = sampleDataKClass.memberProperties
          .associate { it.name to it.findAnnotation<ExcelWriterColumn>() }
        val excelWriterSampleDtoConstructorNamesInOrder =
          sampleDataKClass.constructors.flatMap { constructor ->
            constructor.parameters
          }.mapNotNull { parameter ->
            val excelWriterColumn = memberPropertiesMap[parameter.name]
            val headerName = excelWriterColumn?.headerName
            headerName?.ifBlank { parameter.name }
          }

        info { "${sampleDataKClass.simpleName} constructor names in order: $excelWriterSampleDtoConstructorNamesInOrder" }

        val headerRowCellValues = (0 until headerRow.physicalNumberOfCells).map {
          headerRow.getCell(it).stringCellValue
        }

        info { "Excel file header row cell values: $headerRowCellValues" }

        (0 until headerRow.physicalNumberOfCells).forEach {
          headerRowCellValues[it] shouldBe excelWriterSampleDtoConstructorNamesInOrder[it]
        }
      }
    }

    `when`("headerName is not provided in annotation") {
      then("member's property name is replaced instead") {
        val memberNamesWithoutHeaderNameAnnotated = sampleDataKClass.memberProperties.filter {
          val excelWriterAnnotation = it.findAnnotation<ExcelWriterColumn>()
          excelWriterAnnotation != null && excelWriterAnnotation.headerName.isBlank()
        }.map { it.name }

        info { "Members without header name annotated: $memberNamesWithoutHeaderNameAnnotated" }

        val headerRowCellValues = (0 until headerRow.physicalNumberOfCells).map { columnIdx ->
          headerRow.getCell(columnIdx).stringCellValue
        }

        info { "Excel header row cell values: $headerRowCellValues" }

        headerRowCellValues.containsAll(memberNamesWithoutHeaderNameAnnotated)
      }
    }
  }
})
