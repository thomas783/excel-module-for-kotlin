package excel.writer.annotation

/**
 * Annotation for Excel writer freeze pane option
 * @property colSplit column size to split from left. Default is 0
 * @property rowSplit row size to split from top. Default is 0
 * @see org.apache.poi.ss.usermodel.Sheet.createFreezePane
 */
@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.CLASS)
annotation class ExcelWriterFreezePane(
  val colSplit: Int = 0,
  val rowSplit: Int = 0,
)
