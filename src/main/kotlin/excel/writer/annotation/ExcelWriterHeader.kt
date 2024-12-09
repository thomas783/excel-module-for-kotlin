package excel.writer.annotation

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.CLASS)
annotation class ExcelWriterHeader(
  val essentialFields: Array<String> = [],
)
