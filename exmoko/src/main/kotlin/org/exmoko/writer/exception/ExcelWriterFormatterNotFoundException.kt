package org.exmoko.writer.exception

import kotlin.reflect.KClass

class ExcelWriterFormatterNotFoundException(
  kClass: KClass<*>,
  message: String = "${kClass.simpleName} is not registered in the formatter."
) : ExcelWriterException(message)
