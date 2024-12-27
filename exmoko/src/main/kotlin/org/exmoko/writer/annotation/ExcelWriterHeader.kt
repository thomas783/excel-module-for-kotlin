package org.exmoko.writer.annotation

import org.apache.poi.ss.usermodel.IndexedColors

/**
 * Annotation for Excel writer header options
 * @property name Customized headerName for Excel column.
 * If not provided, it will use the property name itself
 * @property cellColor Customized header cell color. Default [IndexedColors.WHITE]
 * @see IndexedColors
 */
@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.PROPERTY)
annotation class ExcelWriterHeader(
  val name: String = "",
  val cellColor: IndexedColors = IndexedColors.WHITE,
)
