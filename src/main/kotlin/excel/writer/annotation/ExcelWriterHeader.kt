package excel.writer.annotation

import org.apache.poi.ss.usermodel.IndexedColors

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.PROPERTY)
annotation class ExcelWriterHeader(
  val name: String = "",
  val cellColor: IndexedColors = IndexedColors.WHITE,
)
