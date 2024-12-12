package excel.writer

import excel.writer.annotation.ExcelWriterColumn
import excel.writer.annotation.ExcelWriterColumn.Companion.getValidationErrorText
import excel.writer.annotation.ExcelWriterColumn.Companion.getValidationFormula
import excel.writer.annotation.ExcelWriterColumn.Companion.getValidationList
import excel.writer.annotation.ExcelWriterColumn.Companion.getValidationPromptText
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.DataValidationConstraint
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.util.CellRangeAddressList
import org.apache.poi.xssf.streaming.SXSSFCell
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import java.time.LocalDate
import java.time.LocalDateTime
import kotlin.reflect.KClass
import kotlin.reflect.KParameter
import kotlin.reflect.KProperty1
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.isSubclassOf
import kotlin.reflect.full.memberProperties
import kotlin.reflect.jvm.jvmErasure

class ExcelWriter {
  companion object {
    // not to create too much cell styles
    private lateinit var columnCellStyleMap: MutableMap<KClass<*>, CellStyle>
    private val defaultKClasses: List<KClass<*>> = listOf(
      String::class,
      Int::class,
      Double::class,
      Long::class,
      LocalDate::class,
      LocalDateTime::class,
    )

    fun SXSSFWorkbook.initColumnCellStyleMap() {
      columnCellStyleMap = mutableMapOf()
      defaultKClasses.forEach { kClass ->
        val dataFormat = createDataFormat()
        val columnCellStyle = createCellStyle().apply {
          this.dataFormat = getDataFormatByKClass(dataFormat, kClass)
        }
        columnCellStyleMap[kClass] = columnCellStyle
      }
    }

    private fun getDataFormatByKClass(format: DataFormat, kClass: KClass<*>): Short {
      return when (kClass) {
        String::class -> format.getFormat("@")
        Int::class, Long::class -> format.getFormat("0")
        Double::class -> format.getFormat("0.0")
        LocalDate::class -> format.getFormat("yyyy-mm-dd")
        LocalDateTime::class -> format.getFormat("yyyy-mm-dd hh:mm:ss")
        else -> format.getFormat("@")
      }
    }

    inline fun <reified T : Any> createWorkbook(data: Collection<T>, sheetName: String): SXSSFWorkbook {
      val memberProperties = T::class.memberProperties.filter {
        it.findAnnotation<ExcelWriterColumn>() != null
      }
      val parameters: List<KProperty1<T, *>> = T::class.constructors.map { constructor ->
        constructor.parameters.mapNotNull { kParameter: KParameter ->
          memberProperties.firstOrNull { p -> p.name == kParameter.name }
        }
      }.flatten()
      val workbook = SXSSFWorkbook().apply {
        initColumnCellStyleMap()
        createSheet(sheetName).apply {
          // tracking all columns for auto-sizing takes to long
          untrackAllColumnsForAutoSizing()
          setHeaderRows(parameters)
          setValidationConstraints(parameters, data.size)
          setBodyData(data, parameters)
        }
      }

      return workbook
    }

    fun <T : Any> SXSSFSheet.setHeaderRows(kProperties: List<KProperty1<T, *>>) {
      val headerRow = createRow(0)
      kProperties.forEachIndexed { columnIndex, property ->
        val columnAnnotation = property.findAnnotation<ExcelWriterColumn>() ?: return@forEachIndexed
        val headerName = columnAnnotation.headerName.takeIf { it.isNotBlank() } ?: property.name
        val headerCellStyle = createHeaderCellStyle(workbook, columnAnnotation)

        headerRow.createCell(columnIndex).apply {
          setCellValue(headerName)
          cellStyle = headerCellStyle
        }
        setHeaderPromptBox(property, columnIndex)
      }
    }

    private fun createHeaderCellStyle(workbook: SXSSFWorkbook, excelColumn: ExcelWriterColumn): XSSFCellStyle {
      val fontStyle = workbook.createFont().apply {
        bold = true
        fontHeightInPoints = 16
      }
      return workbook.createCellStyle().apply {
        alignment = HorizontalAlignment.CENTER
        fillForegroundColor = excelColumn.headerCellColor.index
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

    fun <T : Any> SXSSFSheet.setBodyData(inputData: Collection<T>, kProperties: List<KProperty1<T, *>>) {
      inputData.mapIndexed { rowIndex, item ->
        val row = this.createRow(rowIndex + 1)
        kProperties.forEachIndexed { columnIndex, property ->
          val cell = row.createCell(columnIndex)
          val value = property.get(item)

          if (value != null) cell.setValueAndDataFormat(property, value)
        }
      }
    }

    private fun <T> SXSSFCell.setValueAndDataFormat(property: KProperty1<T, *>, value: Any) {
      val propertyReturnType = property.returnType.jvmErasure

      this.apply {
        if (propertyReturnType.isSubclassOf(Enum::class)) {
          val enumValue = (value as? Enum<*>)?.name ?: ""
          setCellValue(enumValue)
          cellStyle = columnCellStyleMap[String::class]
        } else {
          when (propertyReturnType) {
            String::class -> setCellValue(value as String)
            LocalDate::class -> setCellValue(value as LocalDate)
            LocalDateTime::class -> setCellValue(value as LocalDateTime)
            Double::class -> setCellValue(value as Double)
            Int::class, Long::class -> setCellValue(value.toString().toDouble())
            else -> setCellValue(value.toString())
          }
          cellStyle = columnCellStyleMap[propertyReturnType]
        }
      }
    }

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
            val constraint = with(excelColumn) {
              helper.createTextLengthConstraint(operationType, operationFormula1, operationFormula2)
            }
            val addressList = CellRangeAddressList(1, dataSize, columnIndex, columnIndex)
            setValidationConstraint(excelColumn, addressList, constraint)
          }

          DataValidationConstraint.ValidationType.DECIMAL -> {
            val constraint = with(excelColumn) {
              helper.createDecimalConstraint(operationType, operationFormula1, operationFormula2)
            }
            val addressList = CellRangeAddressList(1, dataSize, columnIndex, columnIndex)
            setValidationConstraint(excelColumn, addressList, constraint)
          }

          DataValidationConstraint.ValidationType.INTEGER -> {
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
