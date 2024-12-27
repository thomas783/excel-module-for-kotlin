package org.exmoko.writer.annotation

import kotlin.reflect.full.memberProperties

@Retention(AnnotationRetention.RUNTIME)
@Target(AnnotationTarget.CLASS)
annotation class ExcelWritable(
  val properties: Array<String> = [],
) {
  companion object {
    inline fun <reified T : Any> ExcelWritable.getProperties(): Collection<String> {
      return if (properties.isEmpty()) T::class.memberProperties.map { it.name }
      else properties.toList()
    }
  }
}
