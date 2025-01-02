package org.exmoko.writer.annotation

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.PROPERTY)
annotation class ExcelWriterFormat(
  val pattern: String = ""
)
