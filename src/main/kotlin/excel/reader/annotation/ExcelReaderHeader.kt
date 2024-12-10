package excel.reader.annotation

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.CLASS)
annotation class ExcelReaderHeader(
  val essentialFields: Array<String> = [],
)
