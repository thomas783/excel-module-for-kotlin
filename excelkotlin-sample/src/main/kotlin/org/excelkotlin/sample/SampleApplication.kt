package org.excelkotlin.sample

import org.excelkotlin.reader.ExcelReader
import org.excelkotlin.writer.ExcelWriter
import org.excelkotlin.writer.annotation.ExcelWritable
import java.io.File
import java.io.FileOutputStream
import java.time.LocalDate
import java.time.LocalDateTime

@ExcelWritable
data class SampleDto(
  var name: String? = null,
  var age: Int? = null,
  var email: String? = null,
  var address: String? = null,
  var phone: String? = null,
  var birth: LocalDate? = null,
  var lastLoginAt: LocalDateTime? = null
) {
  companion object {
    fun create(size: Int) = (1..size).map {
      SampleDto(
        name = "name$it",
        age = it,
        email = "email$it",
        address = "address$it",
        phone = "phone$it",
        birth = LocalDate.now().minusDays(it.toLong()),
        lastLoginAt = LocalDateTime.now().minusSeconds(it.toLong())
      )
    }
  }
}

fun main() {
  val localPath = "excelkotlin-sample/src/main/resources/sample.xlsx"
  val sampleDataToWrite = SampleDto.create(100)
  val workbook = ExcelWriter.createWorkbook(sampleDataToWrite, "sample")
  val out = FileOutputStream(localPath)

  with(workbook) {
    write(out)
    out.close()
    dispose()
    close()
  }
  val sampleDataAfterRead = ExcelReader(localPath).readExcelFile<SampleDto>()

  sampleDataAfterRead.forEach { println(it) }

  File(localPath).delete()
}
