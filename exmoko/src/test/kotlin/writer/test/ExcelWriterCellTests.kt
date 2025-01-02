package writer.test

import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import io.kotest.matchers.collections.shouldBeIn
import org.apache.poi.ss.usermodel.CellType
import org.exmoko.writer.ExcelWriterFormatter
import org.exmoko.writer.annotation.ExcelWritable
import org.exmoko.writer.annotation.ExcelWritable.Companion.getProperties
import org.exmoko.writer.annotation.ExcelWriterFormat
import writer.dto.ExcelWriterSampleDto
import writer.test.ExcelWriterBaseTests.Companion.setExcelWriterCommonSpec
import java.time.LocalDate
import java.time.LocalDateTime
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.hasAnnotation
import kotlin.reflect.full.isSubclassOf
import kotlin.reflect.full.memberProperties
import kotlin.reflect.jvm.jvmErasure

@OptIn(ExperimentalKotest::class)
class ExcelWriterCellTests : BehaviorSpec({
  val sampleDataSize = 1000
  val sampleDtoKClass = ExcelWriterSampleDto::class
  val sampleData = ExcelWriterSampleDto.createSampleData(sampleDataSize)
  val baseTest = setExcelWriterCommonSpec<ExcelWriterSampleDto.Companion, ExcelWriterSampleDto>(
    sampleDataSize = sampleDataSize,
    path = "sample-cell-value-type-check",
  )

  given("ExcelWritable Annotation") {
    val sheet = baseTest.workbook.getSheetAt(0)
    val excelWritableProperties = sampleDtoKClass.findAnnotation<ExcelWritable>()?.getProperties<ExcelWriterSampleDto>()
      ?.toList()!!
    val sampleDtoConstructorParameters = sampleDtoKClass.constructors.flatMap { it.parameters }
    val sampleDtoConstructorReturnTypeInOrder = sampleDtoConstructorParameters.filter { parameter ->
      parameter.name in excelWritableProperties
    }.map { Triple(it.name, it.type.jvmErasure, it.type.isMarkedNullable) }

    then("excel cell type is set to expected type") {
      sampleDtoConstructorReturnTypeInOrder.forEachIndexed { columnIdx, (propertyName, kClass, isMarkedNullable) ->
        (1..sampleDataSize).forEach { rowIdx ->
          val cell = sheet.getRow(rowIdx).getCell(columnIdx)
          val actualCellType = cell.cellType
          val expectedCellType = when {
            kClass.isSubclassOf(Enum::class) -> CellType.STRING
            else -> when (kClass) {
              String::class -> CellType.STRING
              Int::class, Long::class, Double::class, LocalDate::class, LocalDateTime::class -> CellType.NUMERIC
              else -> CellType.STRING
            }
          }
          val expectedCellTypes =
            if (isMarkedNullable) setOf(CellType.BLANK, expectedCellType) else setOf(expectedCellType)

          debug { "rowIdx: $rowIdx, columnIdx: $columnIdx" }
          debug { "Property Name: $propertyName Expected Cell Types: $expectedCellTypes" }
          debug { "Actual Type: $actualCellType" }

          actualCellType shouldBeIn expectedCellTypes
        }
      }
    }

    then("excel cell is set to expected format") {
      val defaultFormatter = ExcelWriterFormatter()
      val memberProperties = sampleDtoKClass.memberProperties
      sampleDtoConstructorReturnTypeInOrder.forEachIndexed { columnIdx, (propertyName, kClass, isMarkedNullable) ->
        val customFormat = memberProperties.firstOrNull { it.name == propertyName}?.findAnnotation<ExcelWriterFormat>()?.pattern
        val expectedFormat = if (!customFormat.isNullOrBlank()) customFormat
        else defaultFormatter.getFormat(kClass)
        val expectedDataFormats = if (isMarkedNullable) setOf("General", expectedFormat)
        else setOf(expectedFormat)

        (1..sampleDataSize).forEach { rowIdx ->
          val cell = sheet.getRow(rowIdx).getCell(columnIdx)
          val cellDataFormat = cell.cellStyle.dataFormatString

          debug { "rowIdx: $rowIdx, columnIdx: $columnIdx" }
          debug { "Property Name: $propertyName, Expected Data Formats: $expectedDataFormats" }
          debug { "Actual Data Format: $cellDataFormat" }

          cellDataFormat shouldBeIn expectedDataFormats
        }
      }
    }

    then("excel cell data is set to expected value") {
      sampleDtoConstructorReturnTypeInOrder.forEachIndexed { columnIdx, (propertyName, kClass, isMarkedNullable) ->
        (1..sampleDataSize).forEach { rowIdx ->
          val cell = sheet.getRow(rowIdx).getCell(columnIdx)
          val expectedValue = sampleData[rowIdx - 1].let { sampleDto ->
            sampleDto.javaClass.kotlin.memberProperties.first { it.name == propertyName }.get(sampleDto)
          }.let {
            if (it is Enum<*>) it.name else it
          }
          val expectedValues = if (isMarkedNullable) setOf("", expectedValue) else setOf(expectedValue)
          val actualValue = when (kClass) {
            String::class -> cell.richStringCellValue.string
            Int::class -> cell.numericCellValue.toInt()
            Long::class -> cell.numericCellValue.toLong()
            Double::class -> cell.numericCellValue
            LocalDate::class -> cell.localDateTimeCellValue.toLocalDate()
            LocalDateTime::class -> cell.localDateTimeCellValue
            else -> cell.stringCellValue
          }

          debug { "rowIdx: $rowIdx, columnIdx: $columnIdx" }
          debug { "PropertyName: $propertyName, Expected cell values: $expectedValues" }
          debug { "Actual cell value: $actualValue" }

          actualValue shouldBeIn expectedValues
        }
      }
    }
  }
})
