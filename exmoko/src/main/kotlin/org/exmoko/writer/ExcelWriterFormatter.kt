package org.exmoko.writer

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.exmoko.writer.annotation.ExcelWriterFormat
import org.exmoko.writer.exception.ExcelWriterFormatterNotFoundException
import java.time.LocalDate
import java.time.LocalDateTime
import kotlin.reflect.KClass
import kotlin.reflect.KProperty1
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties
import kotlin.reflect.jvm.jvmErasure

/**
 * Formatter for Excel Writer
 *
 * @see org.exmoko.writer.ExcelWriter
 */
class ExcelWriterFormatter {
  private var formatMap: MutableMap<KClass<*>, String> = defaultFormatMap.toMutableMap()
  private lateinit var columnCellStyleMap: MutableMap<KClass<*>, CellStyle>
  private lateinit var customColumnCellStyleMap: MutableMap<String, CellStyle>

  fun initFormatter(workbook: SXSSFWorkbook, kClass: KClass<*>) {
    columnCellStyleMap = mutableMapOf()
    formatMap.entries.forEach { (kClass, format) ->
      with(workbook) {
        val dataFormat = createDataFormat()
        val columnCellStyle = createCellStyle().apply {
          this.dataFormat = dataFormat.getFormat(format)
        }
        columnCellStyleMap[kClass] = columnCellStyle
      }
    }

    customColumnCellStyleMap = mutableMapOf()
    kClass.memberProperties.forEach { member ->
      member.findAnnotation<ExcelWriterFormat>()?.let {
        if (it.pattern.isBlank()) return@forEach
        with(workbook) {
          val dataFormat = createDataFormat()
          val columnCellStyle = createCellStyle().apply {
            this.dataFormat = dataFormat.getFormat(it.pattern)
          }
          customColumnCellStyleMap[member.name] = columnCellStyle
        }
      }
    }
  }

  /**
   * Function to get the format of the cell.
   *
   * @param property [KProperty1] The class of the cell
   * @return [CellStyle] The format of the cell
   * @throws ExcelWriterFormatterNotFoundException If the class is not registered in the formatter
   * @see org.apache.poi.ss.usermodel.DataFormat
   * @see org.apache.poi.ss.usermodel.BuiltinFormats
   */
  fun <T> getCellStyle(property: KProperty1<T, *>): CellStyle {
    val kClass = property.returnType.jvmErasure
    return customColumnCellStyleMap[property.name]
      ?: columnCellStyleMap[kClass]
      ?: columnCellStyleMap[String::class]
      ?: throw ExcelWriterFormatterNotFoundException(kClass)
  }

  fun getFormat(kClass: KClass<*>): String {
    return formatMap[kClass] ?: DEFAULT_FORMAT
  }

  companion object {
    private const val DEFAULT_FORMAT = "@"
    private val defaultFormatMap: Map<KClass<*>, String> = mapOf(
      String::class to DEFAULT_FORMAT,
      Int::class to "0",
      Long::class to "0",
      Double::class to "0.0",
      LocalDate::class to "yyyy-mm-dd",
      LocalDateTime::class to "yyyy-mm-dd hh:mm:ss"
    )
  }
}
