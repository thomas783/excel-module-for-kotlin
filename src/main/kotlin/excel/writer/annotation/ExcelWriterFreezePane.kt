package excel.writer.annotation

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.CLASS)
annotation class ExcelWriterFreezePane(
  val colSplit: Int = 0,
  val rowSplit: Int = 0,
)
