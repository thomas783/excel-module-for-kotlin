package excel.writer.annotation

import excel.writer.exception.ExcelWriterException
import excel.writer.exception.ExcelWriterValidationFormulaException
import excel.writer.exception.ExcelWriterValidationListException
import org.apache.poi.ss.usermodel.DataValidation
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.IndexedColors
import kotlin.reflect.KClass

/**
 * Annotation for Excel writer column options
 * @property headerName Customized headerName for Excel column. If not provided, it will use the property name itself
 * @property headerCellColor Customized header cell color. Default [IndexedColors.WHITE]
 * @property validationType [DataValidationConstraint.ValidationType].
 * Default [DataValidationConstraint.ValidationType.ANY]
 * @property operationType [DataValidationConstraint.OperatorType]. Default [DataValidationConstraint.OperatorType.IGNORED] -1.
 * @property operationFormula1 Customized operation formula 1
 * @property operationFormula2 Customized operation formula 2
 * @property validationIgnoreBlank Ignore blank cell for validation. Use for nullable fields. Default true
 * @property validationListOptions Array of validation list options
 * @property validationListEnum Enum class for validation list options
 * @property validationPromptTitle Title for validation if error occurs
 * @property validationPromptText Text for validation if error occurs
 * @property validationFormula Customized validation formula
 * @property validationErrorStyle Error style for validation [STOP, WARNING, INFO]
 * @property validationErrorTitle Title for validation if error occurs
 * @property validationErrorText Text for validation if error occurs
 * @throws ExcelWriterException if validationListOptions or validationListEnum is not provided
 * @see IndexedColors
 */

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.PROPERTY)
annotation class ExcelWriterColumn(
//  val headerName: String = "",
//  val headerCellColor: IndexedColors = IndexedColors.WHITE,
  val validationType: Int = DataValidationConstraint.ValidationType.ANY,
  val operationType: Int = DEFAULT_OPERATION_TYPE,
  val operationFormula1: String = "",
  val operationFormula2: String = "",
  val validationIgnoreBlank: Boolean = true,
  val validationListOptions: Array<String> = [],
  val validationListEnum: KClass<out Enum<*>> = DefaultValidationListEnum::class,
  val validationPromptTitle: String = "",
  val validationPromptText: String = "",
  val validationFormula: String = "",
  val validationErrorStyle: Int = DataValidation.ErrorStyle.WARNING,
  val validationErrorTitle: String = "",
  val validationErrorText: String = "",
) {
  enum class DefaultValidationListEnum

  companion object {
    fun ExcelWriterColumn.getValidationList(): Array<String> {
      return when {
        validationListOptions.isNotEmpty() -> validationListOptions
        validationListEnum != DefaultValidationListEnum::class -> validationListEnum.java.enumConstants.map { it.name }
          .toTypedArray()

        else -> throw ExcelWriterValidationListException()
      }
    }

    fun ExcelWriterColumn.getValidationFormula(columnIdx: Int, rowIdx: Int): String {
      if (this.validationFormula.isBlank())
        throw ExcelWriterValidationFormulaException()

      return if (this.validationFormula.contains(CURRENT_CELL))
        this.validationFormula.replace(CURRENT_CELL, "${getExcelColumnLetter(columnIdx)}${rowIdx + 1}")
      else this.validationFormula
    }

    fun ExcelWriterColumn.getValidationErrorText(): String {
      return with(this) {
        when {
          validationType == DataValidationConstraint.ValidationType.LIST -> {
            val options = when {
              validationListOptions.isNotEmpty() -> validationListOptions
              validationListEnum != DefaultValidationListEnum::class -> validationListEnum.java.enumConstants
              else -> throw ExcelWriterValidationListException()
            }
            "One of the following values is required. " + options.joinToString(", ")
          }

          else -> validationErrorText
        }
      }
    }

    fun ExcelWriterColumn.getValidationPromptText(): String {
      return with(this) {
        when {
          validationPromptText.isNotBlank() -> validationPromptText
          getValidationErrorText().isNotBlank() -> getValidationErrorText()
          validationPromptTitle.isNotBlank() -> validationPromptTitle
          else -> this.getValidationErrorText()
        }
      }
    }

    private fun getExcelColumnLetter(columnIdx: Int): String {
      var index = columnIdx
      val columnLetter = StringBuilder()

      while (index >= 0) {
        val remainder = index % 26
        columnLetter.insert(0, 'A' + remainder)
        index = (index / 26) - 1
      }

      return columnLetter.toString()
    }

    const val CURRENT_CELL = "CURRENT_CELL"

    const val DEFAULT_OPERATION_TYPE = DataValidationConstraint.OperatorType.IGNORED - 1
  }
}
