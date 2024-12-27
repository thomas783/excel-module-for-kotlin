package reader.test

import org.excelkotlin.reader.ExcelReader
import org.excelkotlin.reader.annotation.ExcelReaderHeader
import org.excelkotlin.reader.exception.ExcelReaderFileExtensionException
import org.excelkotlin.reader.exception.ExcelReaderMissingEssentialHeaderException
import io.kotest.assertions.throwables.shouldThrow
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import io.kotest.matchers.collections.shouldContainAll
import io.kotest.matchers.ints.shouldBeGreaterThan
import io.kotest.matchers.shouldBe
import reader.dto.ExcelReaderSampleDto
import shared.getLocalPath
import kotlin.reflect.full.findAnnotation

@OptIn(ExperimentalKotest::class)
class ExcelReaderBaseTests : BehaviorSpec({
  lateinit var excelReader: ExcelReader

  afterTest {
    excelReader.close()
  }

  given("Excel file for basic tests") {
    val localPath = getLocalPath("sample-to-read.xlsx")
    excelReader = ExcelReader(localPath)
    val excelData = excelReader.readExcelFile<ExcelReaderSampleDto>()

    then("Excel file is read successfully") {
      debug { "Excel data: $excelData" }

      excelData.size shouldBeGreaterThan 0
    }

    `when`("some rows containing cells are all blank") {
      then("should not read those rows") {
        val sheet = excelReader.workbook.getSheetAt(0)
        val physicalNumberOfRows = sheet.physicalNumberOfRows
        val emptyRowSize = (1 until physicalNumberOfRows).map { rowIdx ->
          sheet.getRow(rowIdx)
        }.filter { row ->
          excelReader.isRowAllBlank(row)
        }.size

        debug { "excel row size: $physicalNumberOfRows" }
        debug { "empty row size: ${emptyRowSize}" }
        debug { "excel dto size: ${excelData.size}" }

        // -1 for header row
        excelData.size shouldBe physicalNumberOfRows - emptyRowSize - 1
      }
    }

    `when`("ExcelReaderHeader annotation essential fields are not empty") {
      then("file header should be containing all columns in essential fields") {
        val headerRow = excelReader.getHeader<ExcelReaderSampleDto>().map { it.key }
        val essentialFields = ExcelReaderSampleDto::class.findAnnotation<ExcelReaderHeader>()!!.essentialFields.toList()

        debug { "headerRow: $headerRow" }
        debug { "essentialFieldNames: $essentialFields" }

        headerRow.shouldContainAll(essentialFields)
      }
    }
  }

  given("Excel file for wrong extension") {
    val localPath = getLocalPath("sample-to-read.csv")

    then("ExcelReaderFileExtensionException is thrown") {
      shouldThrow<ExcelReaderFileExtensionException> {
        debug { "path: $localPath" }

        ExcelReader(localPath)
      }.also { debug { it } }
    }
  }

  given("Excel file for missing essential header row") {
    val localPath = getLocalPath("sample-missing-essential-header.xlsx")

    then("ExcelReaderMissingEssentialHeaderException is thrown") {
      shouldThrow<ExcelReaderMissingEssentialHeaderException> {
        ExcelReader(localPath).readExcelFile<ExcelReaderSampleDto>()
      }.also { debug { it } }
    }
  }
})
