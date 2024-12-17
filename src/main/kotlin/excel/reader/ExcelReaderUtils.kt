package excel.reader

import com.github.drapostolos.typeparser.TypeParser
import excel.reader.exception.ExcelReaderInvalidCellTypeException
import org.apache.poi.ss.formula.eval.ErrorEval
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter
import java.util.*
import kotlin.reflect.KProperty1
import kotlin.reflect.jvm.jvmErasure

class ExcelReaderUtils {
  companion object {

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

    fun getValue(cell: Cell?): String {
      if (cell == null || Objects.isNull(cell.cellType)) return ""

      return when (cell.cellType) {
        CellType.STRING -> cell.richStringCellValue.string
        CellType.FORMULA -> {
          runCatching { cell.richStringCellValue.string }.getOrNull()
            ?: runCatching { cell.numericCellValue.toString() }.getOrNull()
            ?: ""
        }

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
  }
}
