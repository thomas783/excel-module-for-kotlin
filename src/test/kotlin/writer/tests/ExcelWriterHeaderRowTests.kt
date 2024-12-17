package writer.tests

import excel.writer.annotation.ExcelWriterColumn
import excel.writer.annotation.ExcelWriterFreezePane
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
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

  given("ExcelWriterFreezePane Annotation") {
    val sheet = baseTest.workbook.getSheetAt(0)

    then("freeze pane is set to annotated row and column") {
      val expectedFreezePane = ExcelWriterSampleDto::class.findAnnotation<ExcelWriterFreezePane>()
      val actualFreezePane = sheet.paneInformation

      debug { "Expected Freeze Pane - Row: ${expectedFreezePane?.rowSplit}, Column: ${expectedFreezePane?.colSplit}" }
      debug { "Actual Freeze Pane - Row: ${actualFreezePane.horizontalSplitTopRow}, Column: ${actualFreezePane.verticalSplitLeftColumn}" }

      expectedFreezePane?.rowSplit shouldBe actualFreezePane.horizontalSplitTopRow
      expectedFreezePane?.colSplit shouldBe actualFreezePane.verticalSplitLeftColumn
    }
  }

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

    given("annotation is provided in constructor") {
      then("header row cell counts equal to ExcelWriterSampleDto properties counts that has ExcelWriterColumn annotation") {
        val excelWriterSampleDtoPropertiesCounts = sampleDataKClass.memberProperties.filter {
          it.hasAnnotation<ExcelWriterColumn>()
        }.size

        debug { "${sampleDataKClass.simpleName} ExcelWriterColumn annotation provided properties count: $excelWriterSampleDtoPropertiesCounts" }
        debug { "Excel file header row cell count: ${headerRow.physicalNumberOfCells}" }

        headerRow.physicalNumberOfCells shouldBe excelWriterSampleDtoPropertiesCounts
      }

      then("header row cell values well created as constructors in order") {
        val sampleDtoHeaderNamesInOrder = sampleDtoConstructorParameters.mapNotNull { parameter ->
          val excelWriterColumn = sampleDtoMemberPropertiesMap[parameter.name]
          val headerName = excelWriterColumn?.headerName
          headerName?.ifBlank { parameter.name }
        }

        debug { "${sampleDataKClass.simpleName} constructor header names in order: $sampleDtoHeaderNamesInOrder" }
        debug { "Excel file header row cell values: $headerRowCellValues" }

        (0 until headerRow.physicalNumberOfCells).forEach { columnIdx ->
          headerRowCellValues[columnIdx] shouldBe sampleDtoHeaderNamesInOrder[columnIdx]
        }
      }
    }

    given("headerName is annotated") {
      then("header row cell values are set to annotated headerName") {
        val memberNamesWithHeaderNameAnnotated = sampleDataKClass.memberProperties.mapNotNull {
          it.findAnnotation<ExcelWriterColumn>()?.headerName
        }.filter { it.isNotBlank() }

        debug { "Members with header name annotated: $memberNamesWithHeaderNameAnnotated" }
        debug { "Excel header row cell values: $headerRowCellValues" }

        headerRowCellValues.containsAll(memberNamesWithHeaderNameAnnotated) shouldBe true
      }
    }

    given("headerName is not annotated") {
      then("member's property name is replaced instead") {
        val memberNamesWithoutHeaderNameAnnotated = sampleDataKClass.memberProperties.filter {
          val excelWriterAnnotation = it.findAnnotation<ExcelWriterColumn>()
          excelWriterAnnotation != null && excelWriterAnnotation.headerName.isBlank()
        }.map { it.name }

        debug { "Members without header name annotated: $memberNamesWithoutHeaderNameAnnotated" }
        debug { "Excel header row cell values: $headerRowCellValues" }

        headerRowCellValues.containsAll(memberNamesWithoutHeaderNameAnnotated) shouldBe true
      }
    }

    given("headerCellColor is annotated") {
      then("header row cell style fillForegroundColor is set to annotated color if not annotated set to default IndexedColors.WHITE") {
        val sampleDtoHeaderCellColorsInOrder = sampleDtoConstructorParameters.mapNotNull { parameter ->
          sampleDtoMemberPropertiesMap[parameter.name]?.headerCellColor
        }
        val headerRowCellStyles = (0 until headerRow.physicalNumberOfCells).map { columnIdx ->
          val colorIndex = headerRow.getCell(columnIdx).cellStyle.fillForegroundColor
          IndexedColors.fromInt(colorIndex.toInt())
        }

        debug { "${sampleDataKClass.simpleName} constructor header cell colors in order: $sampleDtoHeaderCellColorsInOrder" }
        debug { "Excel header row cell colors: $headerRowCellStyles" }

        headerRowCellStyles.indices.forEach {
          headerRowCellStyles[it] shouldBe sampleDtoHeaderCellColorsInOrder[it]
        }
      }
    }
  }
})
