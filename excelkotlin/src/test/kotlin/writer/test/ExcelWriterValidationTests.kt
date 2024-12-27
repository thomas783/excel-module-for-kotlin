package writer.test

import org.excelkotlin.writer.ExcelWriter
import org.excelkotlin.writer.annotation.ExcelWriterColumn
import org.excelkotlin.writer.annotation.ExcelWriterColumn.Companion.DEFAULT_OPERATION_TYPE
import org.excelkotlin.writer.annotation.ExcelWriterColumn.Companion.getValidationErrorText
import org.excelkotlin.writer.annotation.ExcelWriterColumn.Companion.getValidationFormula
import org.excelkotlin.writer.annotation.ExcelWriterColumn.Companion.getValidationPromptText
import org.excelkotlin.writer.exception.ExcelWriterValidationDecimalException
import org.excelkotlin.writer.exception.ExcelWriterValidationFormulaException
import org.excelkotlin.writer.exception.ExcelWriterValidationIntegerException
import org.excelkotlin.writer.exception.ExcelWriterValidationListException
import org.excelkotlin.writer.exception.ExcelWriterValidationTextLengthException
import io.kotest.assertions.throwables.shouldThrow
import io.kotest.common.ExperimentalKotest
import io.kotest.core.spec.style.BehaviorSpec
import io.kotest.engine.test.logging.debug
import io.kotest.matchers.shouldBe
import org.apache.poi.ss.usermodel.DataValidationConstraint
import writer.test.ExcelWriterBaseTests.Companion.setExcelWriterCommonSpec
import writer.dto.ExcelWriterSampleDto
import writer.dto.validation.ExcelWriterValidationTypeDecimalErrorDto
import writer.dto.validation.ExcelWriterValidationTypeFormulaErrorDto
import writer.dto.validation.ExcelWriterValidationTypeIntegerErrorDto
import writer.dto.validation.ExcelWriterValidationTypeListErrorDto
import writer.dto.validation.ExcelWriterValidationTypeTextLengthErrorDto
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
    val sampleDtoMemberPropertiesMap =
      sampleDataKClass.memberProperties.filter { it.hasAnnotation<ExcelWriterColumn>() }
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

        debug { "${sampleDataKClass.simpleName} constructor validation prompt titles in order: $sampleDtoValidationPromptTitlesInOrder" }
        debug { "Excel header row cell validation prompt titles: $headerRowCellValidationPromptTitles" }

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
          sheet.dataValidations.filter { it.regions.cellRangeAddresses.first().containsRow(0) }.map { it.promptBoxText }

        debug { "${sampleDataKClass.simpleName} constructor validation prompt texts in order: $sampleDtoValidationPromptTextsInOrder" }
        debug { "Excel header row cell validation prompt texts: $headerRowCellValidationPromptTexts" }

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

          debug { "${sampleDataKClass.simpleName} constructor validation ignore blank values in order: $sampleDtoValidationIgnoreBlankValues" }
          debug { "Excel data validations by column: $dataValidationsByColumn" }

          sampleDtoValidationIgnoreBlankValues.map { (columnIdx, expectedIgnoreBlankValue) ->
            dataValidationsByColumn[columnIdx]?.forEach { ignoreBlankValue ->
              ignoreBlankValue shouldBe expectedIgnoreBlankValue
            }
          }
        }
      }

      given("validationErrorStyle is annotated") {
        val sampleDtoValidationErrorStyleAnnotated =
          sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
            if (parameter !in sampleDtoValidationTypeAnnotated) null
            else columnIdx to sampleDtoMemberPropertiesMap[parameter.name]?.validationErrorStyle
          }.toMap()
        val errorBoxStyles = sampleDtoValidationErrorStyleAnnotated.keys.associateWith { columnIdx ->
          dataValidationsExceptHeaderRow.filter {
            it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
          }.map { it.errorStyle }
        }

        then("errorStyles are set to annotated validationErrorStyles") {
          sampleDtoValidationErrorStyleAnnotated.keys.forEach { columnIdx ->
            val expectedErrorStyle = sampleDtoValidationErrorStyleAnnotated[columnIdx] ?: return@forEach
            val actualErrorStyle = errorBoxStyles[columnIdx]?.first() ?: return@forEach

            debug { "Expected error style: $expectedErrorStyle" }
            debug { "Excel data actual error style: $actualErrorStyle" }

            expectedErrorStyle shouldBe actualErrorStyle
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

        then("errorBoxTitles are set to annotated validationErrorTitle") {
          sampleDtoValidationErrorTitleAnnotated.keys.forEach { columnIdx ->
            val expectedErrorTitle = sampleDtoValidationErrorTitleAnnotated[columnIdx] ?: return@forEach
            val actualErrorTitle = errorBoxTitles[columnIdx]?.first() ?: return@forEach

            debug { "Expected error title: $expectedErrorTitle" }
            debug { "Excel data actual error title: $actualErrorTitle" }

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

        then("errorBoxTexts are set to as expected validationErrorTexts") {
          sampleDtoValidationErrorTextAnnotated.keys.forEach { columnIdx ->
            val expectedErrorText =
              sampleDtoValidationErrorTextAnnotated[columnIdx]?.getValidationErrorText() ?: return@forEach
            val actualErrorText = errorBoxTexts[columnIdx]?.first() ?: return@forEach

            debug { "Expected error text: $expectedErrorText" }
            debug { "Excel data actual error text: $actualErrorText" }

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

              debug { "Validation list options: ${validationListOptions.joinToString(",")}" }
              debug { "Explicit validation list options: ${explicitValidationListOptions.joinToString(",")}" }

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

              debug { "Validation enum list options: ${validationListEnum.joinToString(",")}" }
              debug { "Explicit validation list options: ${explicitValidationListEnum.joinToString(",")}" }

              validationListEnum shouldBe explicitValidationListEnum
            }
          }
        }

        given("if validationListOptions nor validationListEnum is not annotated") {
          val validationListErrorDto = ExcelWriterValidationTypeListErrorDto.createSampleData(size = sampleDataSize)

          then("ExcelWriterValidationListException is thrown") {
            shouldThrow<ExcelWriterValidationListException> {
              ExcelWriter.createWorkbook(validationListErrorDto, "sample-validation-type-list-error")
            }.also { debug { it } }
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

                debug { "Expected formula: $expectedFormula" }
                debug { "Actual formula: $actualFormula" }

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
              ExcelWriter.createWorkbook(validationFormulaErrorDto, "sample-validation-type-formula-error")
            }.also { debug { it } }
          }
        }
      }

      given("validationType is text_length") {
        given("operationType is annotated") {
          val sampleDtoValidationTextLengthAnnotated =
            sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
              val excelWriterColumn = sampleDtoMemberPropertiesMap[parameter.name]
              if (parameter !in sampleDtoValidationTypeAnnotated || excelWriterColumn?.validationType != DataValidationConstraint.ValidationType.TEXT_LENGTH || excelWriterColumn.operationType == DEFAULT_OPERATION_TYPE) null
              else columnIdx to excelWriterColumn
            }.toMap()
          val dataValidationsByColumn = sampleDtoValidationTextLengthAnnotated.keys.associateWith { columnIdx ->
            dataValidationsExceptHeaderRow.filter {
              it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
            }.map { it.validationConstraint }
          }

          then("data validation is set to annotated") {
            sampleDtoValidationTextLengthAnnotated.keys.forEach { columnIdx ->
              val excelWriterColumn = sampleDtoValidationTextLengthAnnotated[columnIdx] ?: return@forEach
              val dataValidation = dataValidationsByColumn[columnIdx]?.first() ?: return@forEach

              debug {
                with(excelWriterColumn) { "Expected data validation - operationType: $operationType, operationFormula1: $operationFormula1, operationFormula2: $operationFormula2" }
              }
              debug {
                with(dataValidation) { "Actual data validation - operationType: $operator, operationFormula1: $formula1, operationFormula2: $formula2" }
              }

              dataValidation.operator shouldBe excelWriterColumn.operationType
              dataValidation.formula1 shouldBe excelWriterColumn.operationFormula1
              dataValidation.formula2 shouldBe excelWriterColumn.operationFormula2
            }
          }
        }

        given("operationType is not annotated") {
          val validationTextLengthErrorDto =
            ExcelWriterValidationTypeTextLengthErrorDto.createSampleData(size = sampleDataSize)

          then("ExcelWriterValidationTextLengthException is thrown") {
            shouldThrow<ExcelWriterValidationTextLengthException> {
              ExcelWriter.createWorkbook(validationTextLengthErrorDto, "sample-validation-type-text-length-error")
            }.also { debug { it } }
          }
        }
      }

      given("validationType is decimal") {
        given("operationType is annotated") {
          val sampleDtoValidationDecimalAnnotated =
            sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
              val excelWriterColumn = sampleDtoMemberPropertiesMap[parameter.name]
              if (parameter !in sampleDtoValidationTypeAnnotated || excelWriterColumn?.validationType != DataValidationConstraint.ValidationType.DECIMAL || excelWriterColumn.operationType == DEFAULT_OPERATION_TYPE) null
              else columnIdx to excelWriterColumn
            }.toMap()
          val dataValidationsByColumn = sampleDtoValidationDecimalAnnotated.keys.associateWith { columnIdx ->
            dataValidationsExceptHeaderRow.filter {
              it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
            }.map { it.validationConstraint }
          }

          then("data validation is set to annotated") {
            sampleDtoValidationDecimalAnnotated.keys.forEach { columnIdx ->
              val excelWriterColumn = sampleDtoValidationDecimalAnnotated[columnIdx] ?: return@forEach
              val dataValidation = dataValidationsByColumn[columnIdx]?.first() ?: return@forEach

              debug {
                with(excelWriterColumn) { "Expected data validation - operationType: $operationType, operationFormula1: $operationFormula1, operationFormula2: $operationFormula2" }
              }
              debug {
                with(dataValidation) { "Actual data validation - operationType: $operator, operationFormula1: $formula1, operationFormula2: $formula2" }
              }

              dataValidation.operator shouldBe excelWriterColumn.operationType
              dataValidation.formula1 shouldBe excelWriterColumn.operationFormula1
              dataValidation.formula2 shouldBe excelWriterColumn.operationFormula2
            }
          }
        }

        given("operationType is not annotated") {
          val validationDecimalErrorDto =
            ExcelWriterValidationTypeDecimalErrorDto.createSampleData(size = sampleDataSize)

          then("ExcelWriterValidationDecimalException is thrown") {
            shouldThrow<ExcelWriterValidationDecimalException> {
              ExcelWriter.createWorkbook(validationDecimalErrorDto, "sample-validation-type-decimal-error")
            }.also { debug { it } }
          }
        }
      }

      given("validationType is integer") {
        given("operationType is annotated") {
          val sampleDtoValidationIntegerAnnotated =
            sampleDtoConstructorParameters.mapIndexedNotNull { columnIdx, parameter ->
              val excelWriterColumn = sampleDtoMemberPropertiesMap[parameter.name]
              if (parameter !in sampleDtoValidationTypeAnnotated || excelWriterColumn?.validationType != DataValidationConstraint.ValidationType.INTEGER || excelWriterColumn.operationType == DEFAULT_OPERATION_TYPE) null
              else columnIdx to excelWriterColumn
            }.toMap()
          val dataValidationsByColumn = sampleDtoValidationIntegerAnnotated.keys.associateWith { columnIdx ->
            dataValidationsExceptHeaderRow.filter {
              it.regions.cellRangeAddresses.first().containsColumn(columnIdx)
            }.map { it.validationConstraint }
          }
          then("data validation is set to annotated") {
            sampleDtoValidationIntegerAnnotated.keys.forEach { columnIdx ->
              val excelWriterColumn = sampleDtoValidationIntegerAnnotated[columnIdx] ?: return@forEach
              val dataValidation = dataValidationsByColumn[columnIdx]?.first() ?: return@forEach

              debug {
                with(excelWriterColumn) { "Expected data validation - operationType: $operationType, operationFormula1: $operationFormula1, operationFormula2: $operationFormula2" }
              }
              debug {
                with(dataValidation) { "Actual data validation - operationType: $operator, operationFormula1: $formula1, operationFormula2: $formula2" }
              }

              dataValidation.operator shouldBe excelWriterColumn.operationType
              dataValidation.formula1 shouldBe excelWriterColumn.operationFormula1
              dataValidation.formula2 shouldBe excelWriterColumn.operationFormula2
            }
          }
        }

        given("operationType is not annotated") {
          val validationIntegerErrorDto =
            ExcelWriterValidationTypeIntegerErrorDto.createSampleData(size = sampleDataSize)

          then("ExcelWriterValidationIntegerException is thrown") {
            shouldThrow<ExcelWriterValidationIntegerException> {
              ExcelWriter.createWorkbook(validationIntegerErrorDto, "sample-validation-type-integer-error")
            }.also { debug { it } }
          }
        }
      }
    }
  }
})
