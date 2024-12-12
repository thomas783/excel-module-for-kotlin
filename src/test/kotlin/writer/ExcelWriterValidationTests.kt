package writer

import excel.writer.annotation.ExcelWriterColumn
import excel.writer.annotation.ExcelWriterColumn.Companion.getValidationPromptText
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.info
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.ExcelWriterBaseTests.Companion.setExcelWriterCommonSpec
import writer.dto.ExcelWriterSampleDto
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.hasAnnotation
import kotlin.reflect.full.memberProperties

@OptIn(ExperimentalKotest::class)
internal class ExcelWriterValidationTests : BehaviorSpec({
  val sampleDataKClass = ExcelWriterSampleDto::class
  val baseTest = setExcelWriterCommonSpec<ExcelWriterSampleDto.Companion, ExcelWriterSampleDto>(
    sampleDataSize = 1000,
    path = "sample-validation",
  )

  given("ExcelWriterColumn Annotation") {
    val sheet = baseTest.workbook.getSheetAt(0)
    val sampleDtoMemberPropertiesMap = sampleDataKClass.memberProperties
      .filter { it.hasAnnotation<ExcelWriterColumn>() }
      .associate { it.name to it.findAnnotation<ExcelWriterColumn>() }
    val sampleDtoConstructorParameters = sampleDataKClass.constructors.flatMap {
      it.parameters
    }

    given("validationPromptTitle is annotated") {
      then("header row validation prompt title is set to annotated title") {
        val sampleDtoValidationPromptTitlesInOrder = sampleDtoConstructorParameters.mapNotNull { parameter ->
          sampleDtoMemberPropertiesMap[parameter.name]?.validationPromptTitle
        }
        val headerRowCellValidationPromptTitles =
          sheet.dataValidations.filter { it.regions.cellRangeAddresses.first().containsRow(0) }
            .map { it.promptBoxTitle }

        info { "${sampleDataKClass.simpleName} constructor validation prompt titles in order: $sampleDtoValidationPromptTitlesInOrder" }
        info { "Excel header row cell validation prompt titles: $headerRowCellValidationPromptTitles" }

        headerRowCellValidationPromptTitles.indices.forEach {
          headerRowCellValidationPromptTitles[it] shouldBe sampleDtoValidationPromptTitlesInOrder[it]
        }
      }
    }

    given("validationPromptText is annotated") {
      then("header row validation prompt text is set to expected text") {
        val sampleDtoValidationPromptTextsInOrder = sampleDtoConstructorParameters.mapNotNull { parameter ->
          sampleDtoMemberPropertiesMap[parameter.name]?.getValidationPromptText()
        }
        val headerRowCellValidationPromptTexts =
          sheet.dataValidations.filter { it.regions.cellRangeAddresses.first().containsRow(0) }
            .map { it.promptBoxText }

        info { "${sampleDataKClass.simpleName} constructor validation prompt texts in order: $sampleDtoValidationPromptTextsInOrder" }
        info { "Excel header row cell validation prompt texts: $headerRowCellValidationPromptTexts" }

        headerRowCellValidationPromptTexts.indices.forEach {
          headerRowCellValidationPromptTexts[it] shouldBe sampleDtoValidationPromptTextsInOrder[it]
        }
      }
    }

    given("validationType is annotated") {
      val sampleDtoValidationTypeAnnotated = sampleDtoConstructorParameters.filter { parameter ->
        val excelWriterColumn = sampleDtoMemberPropertiesMap[parameter.name]
        excelWriterColumn != null && excelWriterColumn.validationType != DataValidationConstraint.ValidationType.ANY
      }
      val dataValidationsExceptHeaderRow = sheet.dataValidations.filterNot {
        it.regions.cellRangeAddresses.first().containsRow(0)
      }

      given("validationIgnoreBlank is annotated") {
        then("empty cell allowed value is set to annotated validation ignore blank value") {
          val sampleDtoValidationIgnoreBlankValues =
            sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
              if (parameter !in sampleDtoValidationTypeAnnotated) null
              else columnIdx to sampleDtoMemberPropertiesMap[parameter.name]?.validationIgnoreBlank
            }.toMap()
          val dataValidationsByColumn = sampleDtoValidationIgnoreBlankValues.keys.associateWith { columnIdx ->
            dataValidationsExceptHeaderRow.filter {
              it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
            }.map { it.emptyCellAllowed }
          }

          info { "${sampleDataKClass.simpleName} constructor validation ignore blank values in order: $sampleDtoValidationIgnoreBlankValues" }
          info { "Excel data validations by column: $dataValidationsByColumn" }

          sampleDtoValidationIgnoreBlankValues.map { (columnIdx, expectedIgnoreBlankValue) ->
            dataValidationsByColumn[columnIdx]?.forEach { ignoreBlankValue ->
              ignoreBlankValue shouldBe expectedIgnoreBlankValue
            }
          }
        }
      }
    }
  }
})
