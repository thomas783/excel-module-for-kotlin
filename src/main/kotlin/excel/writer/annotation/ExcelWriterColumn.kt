package excel.writer.annotation

import excel.writer.exception.ExcelWriterValidationFormulaException
import excel.writer.exception.ExcelWriterValidationListException
import org.apache.poi.ss.usermodel.DataValidation
import org.apache.poi.ss.usermodel.DataValidationConstraint
import kotlin.reflect.KClass

/**
 * Annotation for Excel writer column options
 *
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
 */

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.PROPERTY)
annotation class ExcelWriterColumn(
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
    /**
     *
     * Extension function to get validation list options
     * @return Array of validation list options in [String]
     * @throws ExcelWriterValidationListException when validationListOptions or validationListEnum is not provided
     */
    @Throws(ExcelWriterValidationListException::class)
    fun ExcelWriterColumn.getValidationList(): Array<String> {
      return when {
        validationListOptions.isNotEmpty() -> validationListOptions
        validationListEnum != DefaultValidationListEnum::class -> validationListEnum.java.enumConstants.map { it.name }
          .toTypedArray()

        else -> throw ExcelWriterValidationListException()
      }
    }

    /**
     * Extension function to get validation formula
     * @param columnIdx index of the column
     * @param rowIdx index of the row
     * @return Customized validation formula
     * @throws ExcelWriterValidationFormulaException when validationFormula is not provided
     */
    @Throws(ExcelWriterValidationFormulaException::class)
    fun ExcelWriterColumn.getValidationFormula(columnIdx: Int, rowIdx: Int): String {
      if (this.validationFormula.isBlank())
        throw ExcelWriterValidationFormulaException()

      return if (this.validationFormula.contains(CURRENT_CELL))
        this.validationFormula.replace(CURRENT_CELL, "${getExcelColumnLetter(columnIdx)}${rowIdx + 1}")
      else this.validationFormula
    }

    /**
     * Extension function to get validation error text
     *
     * If validationType is [DataValidationConstraint.ValidationType.LIST] then it will return the text of requiring validation options
     * @return Customized validation error text
     */
    fun ExcelWriterColumn.getValidationErrorText(): String {
      return if (validationType == DataValidationConstraint.ValidationType.LIST) {
        "One of the following values is required. " + getValidationList().joinToString(", ")
      } else validationErrorText
    }

    /**
     * Extension function to get validation prompt text
     *
     * Priority: validationPromptText > validationErrorText > validationPromptTitle
     * @return Customized validation prompt text
     */
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

    /**
     * Default operation type
     *
     * [DataValidationConstraint.OperatorType.IGNORED] is equal to [DataValidationConstraint.OperatorType.BETWEEN]
     * so need to set the default operation type other than [DataValidationConstraint.OperatorType.IGNORED]
     * @see DataValidationConstraint.OperatorType
     */
    const val DEFAULT_OPERATION_TYPE = DataValidationConstraint.OperatorType.IGNORED - 1
  }
}
