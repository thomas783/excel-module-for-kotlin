package writer.test

import com.excelkotlin.writer.annotation.ExcelWritable
import com.excelkotlin.writer.annotation.ExcelWritable.Companion.getProperties
import com.excelkotlin.writer.annotation.ExcelWriterFreezePane
import com.excelkotlin.writer.annotation.ExcelWriterHeader
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.IndexedColors
import writer.test.ExcelWriterBaseTests.Companion.setExcelWriterCommonSpec
import writer.dto.ExcelWriterSampleDto
import writer.dto.ExcelWriterWritablePropertiesEmptyDto
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.hasAnnotation
import kotlin.reflect.full.memberProperties

@OptIn(ExperimentalKotest::class)
internal class ExcelWriterHeaderRowTests : BehaviorSpec({
  val sampleDataSize = 1000
  val sampleDataKClass = ExcelWriterSampleDto::class
  val baseTest = setExcelWriterCommonSpec<ExcelWriterSampleDto.Companion, ExcelWriterSampleDto>(
    sampleDataSize = sampleDataSize,
    path = "sample-header-row",
  )
  val excelWritablePropertiesEmptyDataKClass = ExcelWriterWritablePropertiesEmptyDto::class
  val excelWritablePropertiesEmptyData =
    setExcelWriterCommonSpec<ExcelWriterWritablePropertiesEmptyDto.Companion, ExcelWriterWritablePropertiesEmptyDto>(
      sampleDataSize = sampleDataSize,
      path = "sample-excel-writable-properties-empty",
    )

  given("ExcelWritable Annotation") {
    val sheet = baseTest.workbook.getSheetAt(0)
    val headerRow = sheet.getRow(0)

    given("ExcelWriterFreezePane Annotation") {
      then("freeze pane is set to annotated row and column") {
        val expectedFreezePane = ExcelWriterSampleDto::class.findAnnotation<ExcelWriterFreezePane>()
        val actualFreezePane = sheet.paneInformation

        debug { "Expected Freeze Pane - Row: ${expectedFreezePane?.rowSplit}, Column: ${expectedFreezePane?.colSplit}" }
        debug { "Actual Freeze Pane - Row: ${actualFreezePane.horizontalSplitTopRow}, Column: ${actualFreezePane.verticalSplitLeftColumn}" }

        expectedFreezePane?.rowSplit shouldBe actualFreezePane.horizontalSplitTopRow
        expectedFreezePane?.colSplit shouldBe actualFreezePane.verticalSplitLeftColumn
      }
    }

    val excelWritablePropertiesIfNotEmpty =
      sampleDataKClass.findAnnotation<ExcelWritable>()?.getProperties<ExcelWriterSampleDto>()?.toList()!!

    `when`("ExcelWritable Annotation's properties are not empty") {

      then("header row cell counts equal to ExcelWritable annotation provided properties count") {

        debug { "${sampleDataKClass.simpleName} ExcelWritable annotation provided properties count: ${excelWritablePropertiesIfNotEmpty.size}" }
        debug { "Excel file header row cell count: ${headerRow.physicalNumberOfCells}" }

        headerRow.physicalNumberOfCells shouldBe excelWritablePropertiesIfNotEmpty.size
      }
    }

    `when`("ExcelWritable Annotation's properties are empty") {
      val excelWritablePropertiesIfEmpty = excelWritablePropertiesEmptyDataKClass.findAnnotation<ExcelWritable>()
        ?.getProperties<ExcelWriterWritablePropertiesEmptyDto>()
      val excelWritablePropertiesEmptyHeaderRow = excelWritablePropertiesEmptyData.workbook.getSheetAt(0).getRow(0)

      then("header row cell counts equal to all class's properties count") {

        debug { "${excelWritablePropertiesEmptyDataKClass.simpleName} all properties count: ${excelWritablePropertiesIfEmpty?.size}" }
        debug { "Excel file header row cell count: ${excelWritablePropertiesEmptyHeaderRow.physicalNumberOfCells}" }

        excelWritablePropertiesEmptyHeaderRow.physicalNumberOfCells shouldBe excelWritablePropertiesIfEmpty?.size
      }
    }

    val headerRowCellValues = (0 until headerRow.physicalNumberOfCells).map { columnIdx ->
      headerRow.getCell(columnIdx).stringCellValue
    }

    given("ExcelWriterHeader Annotation") {
      then("header row cell values well created as annotated properties in order") {
        val expectedHeaderRowCellValuesInOrder = excelWritablePropertiesIfNotEmpty.map { propertyName ->
          sampleDataKClass.memberProperties.find { it.name == propertyName }
            ?.findAnnotation<ExcelWriterHeader>()?.name
            ?.let { it.ifBlank { propertyName } } ?: propertyName
        }

        debug { "${sampleDataKClass.simpleName} annotated properties in order: $expectedHeaderRowCellValuesInOrder" }
        debug { "Excel file header row cell values in order: $headerRowCellValues" }

        (0 until headerRow.physicalNumberOfCells).forEach { columnIdx ->
          headerRowCellValues[columnIdx] shouldBe expectedHeaderRowCellValuesInOrder[columnIdx]
        }
      }

      `when`("ExcelWriterHeader name is not blank") {
        val excelWriterHeaderNotBlankMap: MutableMap<Int, String> = mutableMapOf()
        excelWritablePropertiesIfNotEmpty.forEachIndexed { columnIdx, propertyName ->
          val headerName = sampleDataKClass.memberProperties.find { it.name == propertyName }
            ?.findAnnotation<ExcelWriterHeader>()?.name ?: return@forEachIndexed
          if (headerName.isNotBlank()) excelWriterHeaderNotBlankMap[columnIdx] = headerName
        }

        then("excel header cell value is annotated value") {
          debug { "ExcelWriterHeader annotated not blank names: $excelWriterHeaderNotBlankMap" }
          debug { "Excel file header row cell values: $headerRowCellValues" }

          excelWriterHeaderNotBlankMap.forEach { (columnIdx, headerName) ->
            headerRowCellValues[columnIdx] shouldBe headerName
          }
        }
      }

      `when`("ExcelWriterHeader name is blank") {
        val excelWriterHeaderBlankMap: MutableMap<Int, String> = mutableMapOf()
        excelWritablePropertiesIfNotEmpty.forEachIndexed { columnIdx, propertyName ->
          val headerName = sampleDataKClass.memberProperties.find { it.name == propertyName }
            ?.findAnnotation<ExcelWriterHeader>()?.name ?: return@forEachIndexed
          if (headerName.isBlank()) excelWriterHeaderBlankMap[columnIdx] = propertyName
        }

        then("member's property name is replaced instead") {
          debug { "ExcelWriterHeader annotated is blank so replaced names: $excelWriterHeaderBlankMap" }
          debug { "Excel file header row cell values: $headerRowCellValues" }

          excelWriterHeaderBlankMap.forEach { (columnIdx, headerName) ->
            headerRowCellValues[columnIdx] shouldBe headerName
          }
        }
      }

      val headerRowCellColors = (0 until headerRow.physicalNumberOfCells).map { columnIdx ->
        headerRow.getCell(columnIdx).cellStyle.fillForegroundColor.toInt().let(IndexedColors::fromInt)
      }

      `when`("ExcelWriterHeader cellColor is provided") {
        val excelWriterHeaderCellColorMap: MutableMap<Int, IndexedColors> = mutableMapOf()
        excelWritablePropertiesIfNotEmpty.forEachIndexed { columnIdx, propertyName ->
          val cellColor = sampleDataKClass.memberProperties.find { it.name == propertyName }
            ?.findAnnotation<ExcelWriterHeader>()?.cellColor ?: return@forEachIndexed
          excelWriterHeaderCellColorMap[columnIdx] = cellColor
        }

        then("header row cell color is set to annotated cell color") {
          debug { "ExcelWriterHeader annotated cell colors: $excelWriterHeaderCellColorMap" }
          debug { "Excel file header row cell colors: $headerRowCellColors" }

          excelWriterHeaderCellColorMap.forEach { (columnIdx, cellColor) ->
            headerRowCellColors[columnIdx] shouldBe cellColor
          }
        }
      }
    }

    `when`("ExcelWriterHeader Annotation is missing but ExcelWritable Annotation's properties contains property name") {
      val excelWriterHeaderMissingMap: MutableMap<Int, String> = mutableMapOf()
      excelWritablePropertiesIfNotEmpty.forEachIndexed { columnIdx, propertyName ->
        val property = sampleDataKClass.memberProperties.find { it.name == propertyName } ?: return@forEachIndexed
        if (!property.hasAnnotation<ExcelWriterHeader>() && property.name in excelWritablePropertiesIfNotEmpty)
          excelWriterHeaderMissingMap[columnIdx] = propertyName
      }
      then("header row cell value is created as property name") {
        debug { "ExcelWriterHeader not annotated but property name is contained in ExcelWritable Annotation's properties: $excelWriterHeaderMissingMap" }
        debug { "Excel file header row cell values: $headerRowCellValues" }

        excelWriterHeaderMissingMap.forEach { (columnIdx, headerName) ->
          headerRowCellValues[columnIdx] shouldBe headerName
        }
      }
    }
  }
})
