package org.exmoko.writer

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.util.CellRangeAddressList
import org.apache.poi.xssf.streaming.SXSSFCell
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.exmoko.writer.annotation.ExcelWritable
import org.exmoko.writer.annotation.ExcelWritable.Companion.getProperties
import org.exmoko.writer.annotation.ExcelWriterColumn
import org.exmoko.writer.annotation.ExcelWriterColumn.Companion.getValidationErrorText
import org.exmoko.writer.annotation.ExcelWriterColumn.Companion.getValidationFormula
import org.exmoko.writer.annotation.ExcelWriterColumn.Companion.getValidationList
import org.exmoko.writer.annotation.ExcelWriterColumn.Companion.getValidationPromptText
import org.exmoko.writer.annotation.ExcelWriterFreezePane
import org.exmoko.writer.annotation.ExcelWriterHeader
import org.exmoko.writer.exception.ExcelWritableMissingException
import org.exmoko.writer.exception.ExcelWriterValidationDecimalException
import org.exmoko.writer.exception.ExcelWriterValidationIntegerException
import org.exmoko.writer.exception.ExcelWriterValidationTextLengthException
import java.time.LocalDate
import java.time.LocalDateTime
import kotlin.reflect.KParameter
import kotlin.reflect.KProperty1
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.isSubclassOf
import kotlin.reflect.full.memberProperties
import kotlin.reflect.jvm.jvmErasure

/**
 * Excel Writer
 *
 * This class is meant to be used for writing data to an Excel file.
 * @see [ExcelWritable]
 * @see [ExcelWriterColumn]
 * @see [ExcelWriterHeader]
 * @see [ExcelWriterFreezePane]
 */
class ExcelWriter {
  companion object {
    val formatter by lazy { ExcelWriterFormatter() }

    /**
     * Extension function to set freeze pane based on the [ExcelWriterFreezePane] annotation
     * @see ExcelWriterFreezePane
     */
    inline fun <reified T : Any> SXSSFSheet.setFreezePane() {
      T::class.findAnnotation<ExcelWriterFreezePane>()?.let {
        createFreezePane(it.colSplit, it.rowSplit)
      }
    }

    /**
     * Function to create a workbook
     * @param data [Collection]
     * @param sheetName [String] name for the sheet in the workbook
     * @return [SXSSFWorkbook]
     * @throws [ExcelWritableMissingException] If [ExcelWritable] not annotated to given data
     */
    inline fun <reified T : Any> createWorkbook(data: Collection<T>, sheetName: String): SXSSFWorkbook {
      val excelWritableProperties = T::class.findAnnotation<ExcelWritable>()?.getProperties<T>()
        ?: throw ExcelWritableMissingException()

      val memberProperties = T::class.memberProperties.filter {
        it.name in excelWritableProperties
      }
      val parameters: List<KProperty1<T, *>> = T::class.constructors.map { constructor ->
        constructor.parameters.mapNotNull { kParameter: KParameter ->
          memberProperties.firstOrNull { p -> p.name == kParameter.name }
        }
      }.flatten()
      val workbook = SXSSFWorkbook().apply {
        formatter.initFormatter(this, T::class)
        createSheet(sheetName).apply {
          // tracking all columns for auto-sizing takes to long
          untrackAllColumnsForAutoSizing()
          setFreezePane<T>()
          setHeaderRows(parameters)
          setValidationConstraints(parameters, data.size)
          setBodyData(data, parameters)
        }
      }

      return workbook
    }

    /**
     * Extension function to set header rows
     * @param kProperties list of [KProperty1]
     */
    fun <T : Any> SXSSFSheet.setHeaderRows(kProperties: List<KProperty1<T, *>>) {
      val headerRow = createRow(0)
      kProperties.forEachIndexed { columnIndex, property ->
        val columnAnnotation = property.findAnnotation<ExcelWriterHeader>() ?: ExcelWriterHeader()
        val headerName = columnAnnotation.name.takeIf { it.isNotBlank() } ?: property.name
        val headerCellStyle = workbook.createHeaderCellStyle(columnAnnotation.cellColor)

        headerRow.createCell(columnIndex).apply {
          setCellValue(headerName)
          cellStyle = headerCellStyle
        }
        setHeaderPromptBox(property, columnIndex)
      }
    }

    /**
     * Extension function for creating header cell styles
     *
     * If you want to change the default header cell styles, you can change here
     */
    private fun SXSSFWorkbook.createHeaderCellStyle(indexedColors: IndexedColors): XSSFCellStyle {
      val fontStyle = createFont().apply {
        bold = true
        fontHeightInPoints = 16
      }
      return createCellStyle().apply {
        alignment = HorizontalAlignment.CENTER
        fillForegroundColor = indexedColors.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
        borderTop = BorderStyle.THIN
        borderBottom = BorderStyle.THIN
        borderLeft = BorderStyle.THIN
        borderRight = BorderStyle.THIN
        topBorderColor = IndexedColors.BLACK.index
        bottomBorderColor = IndexedColors.BLACK.index
        leftBorderColor = IndexedColors.BLACK.index
        rightBorderColor = IndexedColors.BLACK.index
        setFont(fontStyle)
      } as XSSFCellStyle
    }

    /**
     * Extension function to set header prompt box
     * @param property [KProperty1]
     */
    private fun <T : Any> SXSSFSheet.setHeaderPromptBox(property: KProperty1<T, *>, columnIdx: Int) {
      val excelColumn = property.findAnnotation<ExcelWriterColumn>() ?: return
      val helper = this.dataValidationHelper
      val dummyConstraint =
        helper.createTextLengthConstraint(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "0", null)
      val addressList = CellRangeAddressList(0, 0, columnIdx, columnIdx)
      val validation = helper.createValidation(dummyConstraint, addressList).apply {
        showPromptBox = true
        createPromptBox(excelColumn.validationPromptTitle, excelColumn.getValidationPromptText())
      }

      this.addValidationData(validation)
    }

    /**
     * Extension function to set the cell data
     * @param inputData [Collection]
     * @param kProperties list of [KProperty1]
     */
    fun <T : Any> SXSSFSheet.setBodyData(inputData: Collection<T>, kProperties: List<KProperty1<T, *>>) {
      inputData.mapIndexed { rowIndex, item ->
        val row = this.createRow(rowIndex + 1)
        kProperties.forEachIndexed { columnIndex, property ->
          val cell = row.createCell(columnIndex)
          val value = property.get(item)

          if (value != null) cell.apply {
            setValue(property, value)
            setDataFormat(property)
          }
        }
      }
    }

    /**
     * Extension function to set the cell value and data format
     * @param property [KProperty1]
     * @param value [Any]
     */
    private fun <T> SXSSFCell.setValue(property: KProperty1<T, *>, value: Any) {
      val propertyReturnType = property.returnType.jvmErasure

      if (propertyReturnType.isSubclassOf(Enum::class)) {
        val enumValue = (value as? Enum<*>)?.name ?: ""
        setCellValue(enumValue)
      } else {
        when (propertyReturnType) {
          String::class -> setCellValue(value as String)
          LocalDate::class -> setCellValue(value as LocalDate)
          LocalDateTime::class -> setCellValue(value as LocalDateTime)
          Double::class -> setCellValue(value as Double)
          Int::class, Long::class -> setCellValue(value.toString().toDouble())
          else -> setCellValue(value.toString())
        }
      }
    }

    /**
     * Extension function to set the data format for cell
     *
     * @param property [KProperty1]
     */
    private fun <T> SXSSFCell.setDataFormat(property: KProperty1<T, *>) {
      cellStyle = formatter.getCellStyle(property)
    }

    /**
     * Extension function to set validation constraints
     *
     * If validationType is [DataValidationConstraint.ValidationType.FORMULA] then it will set the formula for each cell
     * else then it will set the validation constraint for the whole column by the data size
     * @param kProperties list of [KProperty1]
     * @param dataSize [Int] size of the data
     */
    fun <T : Any> SXSSFSheet.setValidationConstraints(
      kProperties: List<KProperty1<T, *>>,
      dataSize: Int,
    ) {
      val helper = this.dataValidationHelper
      kProperties.forEachIndexed { columnIndex, property ->
        val excelColumn = property.findAnnotation<ExcelWriterColumn>() ?: return@forEachIndexed
        when (excelColumn.validationType) {
          DataValidationConstraint.ValidationType.FORMULA -> {
            (1..dataSize).forEach { rowIndex ->
              val formula = excelColumn.getValidationFormula(columnIndex, rowIndex)
              val constraint = helper.createCustomConstraint(formula)
              val addressList = CellRangeAddressList(rowIndex, rowIndex, columnIndex, columnIndex)
              setValidationConstraint(excelColumn, addressList, constraint)
            }
          }

          DataValidationConstraint.ValidationType.LIST -> {
            val options = excelColumn.getValidationList()
            val constraint = helper.createExplicitListConstraint(options)
            val addressList = CellRangeAddressList(1, dataSize, columnIndex, columnIndex)
            setValidationConstraint(excelColumn, addressList, constraint)
          }

          DataValidationConstraint.ValidationType.TEXT_LENGTH -> {
            if (excelColumn.operationType == ExcelWriterColumn.DEFAULT_OPERATION_TYPE) {
              throw ExcelWriterValidationTextLengthException()
            }
            val constraint = with(excelColumn) {
              helper.createTextLengthConstraint(operationType, operationFormula1, operationFormula2)
            }
            val addressList = CellRangeAddressList(1, dataSize, columnIndex, columnIndex)
            setValidationConstraint(excelColumn, addressList, constraint)
          }

          DataValidationConstraint.ValidationType.DECIMAL -> {
            if (excelColumn.operationType == ExcelWriterColumn.DEFAULT_OPERATION_TYPE)
              throw ExcelWriterValidationDecimalException()
            val constraint = with(excelColumn) {
              helper.createDecimalConstraint(operationType, operationFormula1, operationFormula2)
            }
            val addressList = CellRangeAddressList(1, dataSize, columnIndex, columnIndex)
            setValidationConstraint(excelColumn, addressList, constraint)
          }

          DataValidationConstraint.ValidationType.INTEGER -> {
            if (excelColumn.operationType == ExcelWriterColumn.DEFAULT_OPERATION_TYPE)
              throw ExcelWriterValidationIntegerException()
            val constraint = with(excelColumn) {
              helper.createIntegerConstraint(operationType, operationFormula1, operationFormula2)
            }
            val addressList = CellRangeAddressList(1, dataSize, columnIndex, columnIndex)
            setValidationConstraint(excelColumn, addressList, constraint)
          }

          else -> return@forEachIndexed
        }
      }
    }

    /**
     * Extension function to set validation constraint
     *
     * @param excelColumn [ExcelWriterColumn]
     * @param addressList [CellRangeAddressList]
     * @param constraint [DataValidationConstraint]
     */
    private fun SXSSFSheet.setValidationConstraint(
      excelColumn: ExcelWriterColumn,
      addressList: CellRangeAddressList,
      constraint: DataValidationConstraint,
    ) {
      val validation = this.dataValidationHelper.createValidation(constraint, addressList).apply {
        showErrorBox = true
        emptyCellAllowed = excelColumn.validationIgnoreBlank
        errorStyle = excelColumn.validationErrorStyle
        with(excelColumn) { createErrorBox(validationErrorTitle, getValidationErrorText()) }
      }

      this.addValidationData(validation)
    }
  }
}
