package excel.writer.annotation

import excel.writer.exception.ExcelWriterException
import org.apache.poi.ss.usermodel.DataValidation
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.IndexedColors
import kotlin.reflect.KClass

/**
 * Excel column annotation for writer
 * @param headerName Customized headerName for Excel column. If not provided, it will use the property name itself
 * @param headerCellColor Customized header cell color. Default [IndexedColors.WHITE]
 * @param validationType [DataValidationConstraint.ValidationType]
 * @param operationType [DataValidationConstraint.OperatorType]
 * @param operationFormula1 Customized operation formula 1
 * @param operationFormula2 Customized operation formula 2
 * @param validationIgnoreBlank Ignore blank cell for validation
 * @param validationListOptions Array of validation list options
 * @param validationListEnum Enum class for validation list options
 * @param validationPromptTitle Title for validation if error occurs
 * @param validationPromptText Text for validation if error occurs
 * @param validationFormula Customized validation formula
 * @param validationErrorStyle Error style for validation [STOP, WARNING, INFO]
 * @param validationErrorTitle Title for validation if error occurs
 * @param validationErrorText Text for validation if error occurs
 * @throws ExcelWriterException if validationListOptions or validationListEnum is not provided
 * @see IndexedColors
 */

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.PROPERTY)
annotation class ExcelWriterColumn(
  val headerName: String = "",
  val headerCellColor: IndexedColors = IndexedColors.WHITE,
  val validationType: Int = DataValidationConstraint.ValidationType.ANY,
  val operationType: Int = DataValidationConstraint.OperatorType.IGNORED,
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

        else -> throw ExcelWriterException("ExcelColumn with either validationListOptions or validationListEnum is required")
      }
    }

    fun ExcelWriterColumn.getValidationFormula(columnIdx: Int, rowIdx: Int): String {
      if (this.validationFormula.isBlank())
        throw ExcelWriterException("ExcelColumn with validationFormula is required")

      return if (this.validationFormula.contains(CURRENT_CELL))
        this.validationFormula.replace(CURRENT_CELL, "${getExcelColumnLetter(columnIdx)}${rowIdx + 1}")
      else this.validationFormula
    }

    fun ExcelWriterColumn.getValidationErrorText(): String {
      return when {
        this.validationListEnum != DefaultValidationListEnum::class
        -> "One of the following values is required. " + this.validationListEnum.java.enumConstants.joinToString(
          ", "
        ) { it.name }

        else -> this.validationErrorText
      }
    }

    fun ExcelWriterColumn.getValidationPromptText(): String {
      return when {
        this.validationPromptText.isNotBlank() -> this.validationPromptText
        this.getValidationErrorText().isNotBlank() -> this.getValidationErrorText()
        this.validationPromptTitle.isNotBlank() -> this.validationPromptTitle
        else -> this.getValidationErrorText()
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
  }
}
