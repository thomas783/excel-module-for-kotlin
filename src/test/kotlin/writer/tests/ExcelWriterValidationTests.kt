package writer.tests

import excel.writer.ExcelWriter
import excel.writer.annotation.ExcelWriterColumn
import excel.writer.annotation.ExcelWriterColumn.Companion.getValidationErrorText
import excel.writer.annotation.ExcelWriterColumn.Companion.getValidationFormula
import excel.writer.annotation.ExcelWriterColumn.Companion.getValidationPromptText
import excel.writer.exception.ExcelWriterValidationFormulaException
import excel.writer.exception.ExcelWriterValidationListException
import io.kotest.assertions.throwables.shouldThrow
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.info
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.DataValidationConstraint
import shared.ExcelWriterBaseTests.Companion.setExcelWriterCommonSpec
import writer.dto.ExcelWriterSampleDto
import writer.dto.ExcelWriterValidationTypeFormulaErrorDto
import writer.dto.ExcelWriterValidationTypeListErrorDto
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.hasAnnotation
import kotlin.reflect.full.memberProperties

@OptIn(ExperimentalKotest::class)
internal class ExcelWriterValidationTests : BehaviorSpec({
  val sampleDataSize = 1000
  val sampleDataKClass = ExcelWriterSampleDto::class
  val baseTest = setExcelWriterCommonSpec<ExcelWriterSampleDto.Companion, ExcelWriterSampleDto>(
    sampleDataSize = sampleDataSize,
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

      given("validationErrorTitle is annotated") {
        val sampleDtoValidationErrorTitleAnnotated =
          sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
            if (parameter !in sampleDtoValidationTypeAnnotated) null
            else columnIdx to sampleDtoMemberPropertiesMap[parameter.name]?.validationErrorTitle
          }.toMap()
        val errorBoxTitles = sampleDtoValidationErrorTitleAnnotated.keys.associateWith { columnIdx ->
          dataValidationsExceptHeaderRow.filter {
            it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
          }.map { it.errorBoxTitle }
        }

        then("errorBoxTitles is set to annotated validationErrorTitle") {
          sampleDtoValidationErrorTitleAnnotated.keys.forEach { columnIdx ->
            val expectedErrorTitle = sampleDtoValidationErrorTitleAnnotated[columnIdx] ?: return@forEach
            val actualErrorTitle = errorBoxTitles[columnIdx]?.first() ?: return@forEach

            info { "Expected error title: $expectedErrorTitle" }
            info { "Excel data actual error title: $actualErrorTitle" }

            expectedErrorTitle shouldBe actualErrorTitle
          }
        }
      }

      given("validationErrorText is annotated") {
        val sampleDtoValidationErrorTextAnnotated =
          sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
            if (parameter !in sampleDtoValidationTypeAnnotated) null
            else columnIdx to sampleDtoMemberPropertiesMap[parameter.name]
          }.toMap()
        val errorBoxTexts = sampleDtoValidationErrorTextAnnotated.keys.associateWith { columnIdx ->
          dataValidationsExceptHeaderRow.filter {
            it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
          }.map { it.errorBoxText }
        }

        then("errorBoxTexts is set to as expected validationErrorTexts") {
          sampleDtoValidationErrorTextAnnotated.keys.forEach { columnIdx ->
            val expectedErrorText =
              sampleDtoValidationErrorTextAnnotated[columnIdx]?.getValidationErrorText() ?: return@forEach
            val actualErrorText = errorBoxTexts[columnIdx]?.first() ?: return@forEach

            info { "Expected error text: $expectedErrorText" }
            info { "Excel data actual error text: $actualErrorText" }

            expectedErrorText shouldBe actualErrorText
          }
        }
      }

      given("validationType is list") {
        given("if validationListOptions is annotated") {
          val sampleDtoValidationListOptionsAnnotated =
            sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
              val validationListOptions = sampleDtoMemberPropertiesMap[parameter.name]?.validationListOptions
              if (parameter !in sampleDtoValidationTypeAnnotated || (validationListOptions != null && validationListOptions.isEmpty())) null
              else columnIdx to sampleDtoMemberPropertiesMap[parameter.name]?.validationListOptions
            }.toMap()
          val dataValidationsByColumn = sampleDtoValidationListOptionsAnnotated.keys.associateWith { columnIdx ->
            dataValidationsExceptHeaderRow.filter {
              it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
            }.map { it.validationConstraint }
          }

          then("explicit data validation is set to annotated validationListOptions") {
            sampleDtoValidationListOptionsAnnotated.keys.forEach { columnIdx ->
              val validationListOptions = sampleDtoValidationListOptionsAnnotated[columnIdx] ?: return@forEach
              val explicitValidationListOptions =
                dataValidationsByColumn[columnIdx]?.first()?.explicitListValues ?: return@forEach

              info { "Validation list options: ${validationListOptions.joinToString(",")}" }
              info { "Explicit validation list options: ${explicitValidationListOptions.joinToString(",")}" }

              validationListOptions shouldBe explicitValidationListOptions
            }
          }
        }

        given("if validationListEnum is annotated") {
          val sampleDtoValidationListEnumAnnotated =
            sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
              val validationListEnum = sampleDtoMemberPropertiesMap[parameter.name]?.validationListEnum
              if (parameter !in sampleDtoValidationTypeAnnotated || (validationListEnum != null && validationListEnum == ExcelWriterColumn.DefaultValidationListEnum::class)) null
              else columnIdx to sampleDtoMemberPropertiesMap[parameter.name]?.validationListEnum?.java?.enumConstants?.map { it.name }
            }.toMap()
          val dataValidationsByColumn = sampleDtoValidationListEnumAnnotated.keys.associateWith { columnIdx ->
            dataValidationsExceptHeaderRow.filter {
              it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
            }.map { it.validationConstraint }
          }

          then("explicit data validation is set to annotated enum constants") {
            sampleDtoValidationListEnumAnnotated.keys.forEach { columnIdx ->
              val validationListEnum = sampleDtoValidationListEnumAnnotated[columnIdx] ?: return@forEach
              val explicitValidationListEnum =
                dataValidationsByColumn[columnIdx]?.first()?.explicitListValues ?: return@forEach

              info { "Validation enum list options: ${validationListEnum.joinToString(",")}" }
              info { "Explicit validation list options: ${explicitValidationListEnum.joinToString(",")}" }

              validationListEnum shouldBe explicitValidationListEnum
            }
          }
        }

        given("if validationListOptions nor validationListEnum is not annotated") {
          val validationListErrorDto =
            ExcelWriterValidationTypeListErrorDto.createSampleData(size = sampleDataSize)

          then("ExcelWriterValidationListException is thrown") {
            shouldThrow<ExcelWriterValidationListException> {
              ExcelWriter.createWorkbook(validationListErrorDto, "sample-validation-list-type-error")
            }.also { info { it } }
          }
        }
      }

      given("validationType is formula") {
        given("validationFormula is annotated") {
          val sampleDtoValidationFormulaAnnotated =
            sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
              val validationFormula = sampleDtoMemberPropertiesMap[parameter.name]?.validationFormula
              if (parameter !in sampleDtoValidationTypeAnnotated || (validationFormula != null && validationFormula.isBlank())) null
              else columnIdx to sampleDtoMemberPropertiesMap[parameter.name]
            }.toMap()
          val dataValidationsByColumn = sampleDtoValidationFormulaAnnotated.keys.associateWith { columnIdx ->
            dataValidationsExceptHeaderRow.filter {
              it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
            }.map { it.validationConstraint }
          }

          then("validationFormula") {
            sampleDtoValidationFormulaAnnotated.keys.forEach { columnIdx ->
              val annotatedValidationFormula = sampleDtoValidationFormulaAnnotated[columnIdx] ?: return@forEach
              val validationFormulas = dataValidationsByColumn[columnIdx] ?: return@forEach

              (1..sampleDataSize).forEach { rowIdx ->
                val expectedFormula = annotatedValidationFormula.getValidationFormula(columnIdx, rowIdx)
                val actualFormula = validationFormulas[rowIdx - 1].formula1

                info { "Expected formula: $expectedFormula" }
                info { "Actual formula: $actualFormula" }

                expectedFormula shouldBe actualFormula
              }
            }
          }
        }

        given("validationFormula is blank") {
          val validationFormulaErrorDto =
            ExcelWriterValidationTypeFormulaErrorDto.createSampleData(size = sampleDataSize)

          then("ExcelWriterValidationFormulaException is thrown") {
            shouldThrow<ExcelWriterValidationFormulaException> {
              ExcelWriter.createWorkbook(validationFormulaErrorDto, "sample-validation-formula-type-error")
            }.also { info { it } }
          }
        }
      }
    }
  }
})
