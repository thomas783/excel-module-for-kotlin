package writer

import excel.writer.annotation.ExcelWriterColumn
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.info
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.IndexedColors
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
    val headerRowCellValues = (0 until headerRow.physicalNumberOfCells).map { columnIdx ->
      headerRow.getCell(columnIdx).stringCellValue
    }
    val sampleDtoMemberPropertiesMap = sampleDataKClass.memberProperties
      .filter { it.hasAnnotation<ExcelWriterColumn>() }
      .associate { it.name to it.findAnnotation<ExcelWriterColumn>() }
    val sampleDtoConstructorParameters = sampleDataKClass.constructors.flatMap {
      it.parameters
    }

    `when`("annotation is provided in constructor") {
      then("header row cell counts equal to ExcelWriterSampleDto properties counts that has ExcelWriterColumn annotation") {
        val excelWriterSampleDtoPropertiesCounts = sampleDataKClass.memberProperties.filter {
          it.hasAnnotation<ExcelWriterColumn>()
        }.size
        headerRow.physicalNumberOfCells shouldBe excelWriterSampleDtoPropertiesCounts
      }

      then("header row cell values well created as constructors in order") {
        val sampleDtoHeaderNamesInOrder = sampleDtoConstructorParameters.mapNotNull { parameter ->
          val excelWriterColumn = sampleDtoMemberPropertiesMap[parameter.name]
          val headerName = excelWriterColumn?.headerName
          headerName?.ifBlank { parameter.name }
        }

        info { "${sampleDataKClass.simpleName} constructor header names in order: $sampleDtoHeaderNamesInOrder" }
        info { "Excel file header row cell values: $headerRowCellValues" }

        (0 until headerRow.physicalNumberOfCells).forEach {
          headerRowCellValues[it] shouldBe sampleDtoHeaderNamesInOrder[it]
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
        info { "Excel header row cell values: $headerRowCellValues" }

        headerRowCellValues.containsAll(memberNamesWithoutHeaderNameAnnotated) shouldBe true
      }
    }

    `when`("headerCellColor is provided") {
      then("header row cell style fillForegroundColor is set to provided color") {
        val sampleDtoHeaderCellColorsInOrder = sampleDtoConstructorParameters.mapNotNull { parameter ->
          sampleDtoMemberPropertiesMap[parameter.name]?.headerCellColor
        }
        val headerRowCellStyles = (0 until headerRow.physicalNumberOfCells).map { columnIdx ->
          val colorIndex = headerRow.getCell(columnIdx).cellStyle.fillForegroundColor
          IndexedColors.fromInt(colorIndex.toInt())
        }

        info { "${sampleDataKClass.simpleName} constructor header cell colors in order: $sampleDtoHeaderCellColorsInOrder" }
        info { "Excel header row cell colors: $headerRowCellStyles" }

        (0 until headerRow.physicalNumberOfCells).forEach {
          headerRowCellStyles[it] shouldBe sampleDtoHeaderCellColorsInOrder[it]
        }
      }
    }
  }
})
