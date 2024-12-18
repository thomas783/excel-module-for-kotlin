package excel.reader

import com.github.drapostolos.typeparser.TypeParser
import com.github.drapostolos.typeparser.TypeParserException
import excel.reader.exception.ExcelReaderInvalidCellTypeException
import org.apache.commons.lang3.StringUtils
import org.apache.commons.lang3.exception.ExceptionUtils
import org.apache.poi.ss.formula.eval.ErrorEval
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Row
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.util.*
import javax.validation.Validation
import javax.validation.ValidationException
import javax.validation.Validator
import kotlin.reflect.KProperty1
import kotlin.reflect.full.createInstance
import kotlin.reflect.jvm.isAccessible
import kotlin.reflect.jvm.javaField
import kotlin.reflect.jvm.jvmErasure

class ExcelReaderUtils {
  companion object {
    private val validator: Validator = Validation.buildDefaultValidatorFactory().validator

    val typeParser: TypeParser = TypeParser.newBuilder()
      .registerParser(LocalDate::class.java) { input, _ ->
        LocalDate.parse(input, DateTimeFormatter.ISO_DATE_TIME)
      }
      .registerParser(LocalDateTime::class.java) { input, _ ->
        LocalDateTime.parse(input, DateTimeFormatter.ISO_DATE_TIME)
      }
      .build()

    fun <T> checkCellType(cell: Cell?, property: KProperty1<T, *>) {
      val cellType = cell?.cellType ?: return

      if (property.returnType.jvmErasure in listOf(LocalDate::class, LocalDateTime::class) &&
        cellType != CellType.NUMERIC
      )
        throw ExcelReaderInvalidCellTypeException("Invalid cell type. The field type must be a date type.")
    }

    fun getValue(cell: Cell?): String? {
      if (cell == null || Objects.isNull(cell.cellType)) return ""

      return when (cell.cellType) {
        CellType.STRING -> cell.richStringCellValue.string
        CellType.FORMULA ->
          runCatching { cell.richStringCellValue.string }.getOrNull()
            ?: runCatching { cell.numericCellValue.toString() }.getOrNull()
            ?: ""

        CellType.NUMERIC -> {
          val value = if (DateUtil.isCellDateFormatted(cell)) cell.localDateTimeCellValue.toString()
          else cell.numericCellValue.toString()
          if (value.endsWith(".0")) value.substring(0, value.length - 2)
          else value
        }

        CellType.BOOLEAN -> cell.booleanCellValue.toString()
        CellType.ERROR -> ErrorEval.getText(cell.errorCellValue.toInt())
        CellType.BLANK, CellType._NONE -> ""
        else -> ""
      }
    }

    inline fun <reified T : Any> setObjectMapping(obj: T, row: Row): T {
      val headerMap = ExcelReader.getHeader<T>()
      val errorFieldList = ExcelReader.excelReaderItem.errorFieldList

      headerMap.mapValues { (_, excelHeaderValue) ->
        val (headerName, headerIdx, field) = excelHeaderValue
        var cellValue: String? = null
        val cell = row.getCell(headerIdx)

        runCatching {
          cellValue = getValue(cell)
          var setData: Any? = null

          if (!cellValue.isNullOrBlank()) checkCellType(cell, field)
          if (!StringUtils.isEmpty(cellValue)) setData = typeParser.parseType(cellValue, field.javaField?.type)
          field.isAccessible = true
          field.setter.call(obj, setData)
          checkValidation(obj, field.name)
        }.onFailure { exception ->
          val (error, message) = when (exception) {
            is ExcelReaderInvalidCellTypeException -> ExcelReaderFieldError.TYPE to ExcelReaderFieldError.TYPE.message
            is TypeParserException -> ExcelReaderFieldError.TYPE to "${exception.message} Field Type: ${field.javaField?.type?.simpleName}, Input Type: ${cellValue?.javaClass?.simpleName}"
            is ValidationException -> ExcelReaderFieldError.VALID to ExcelReaderFieldError.VALID.message
            else -> ExcelReaderFieldError.UNKNOWN to ExcelReaderFieldError.UNKNOWN.message
          }
          errorFieldList.add(
            ExcelReaderErrorField(
              type = error.name,
              row = row.rowNum + 1,
              field = field.name,
              fieldHeader = headerName,
              inputData = cellValue,
              message = message,
              exceptionMessage = ExceptionUtils.getRootCauseMessage(exception)
            )
          )
        }
      }

      return obj
    }

    fun <T> checkValidation(obj: T, fieldName: String) {
      validator.validate(obj)
        .firstOrNull { data -> data.propertyPath.toString() == fieldName }
        ?.run { throw ValidationException() }
    }
  }

  inline fun <reified T : Any> from(row: Row): T = setObjectMapping(T::class.createInstance(), row)
}
