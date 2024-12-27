package org.exmoko.sample

import io.kotest.core.spec.style.ShouldSpec
import io.kotest.matchers.shouldBe
import org.exmoko.reader.ExcelReader
import org.exmoko.writer.ExcelWriter
import java.io.File
import java.io.FileOutputStream

class SampleApplicationTest : ShouldSpec({
  lateinit var excelFile: File
  val sampleDataToWrite = SampleDto.create(100)
  val localPath = "src/test/resources/sample.xlsx"

  beforeSpec {
    val workbook = ExcelWriter.createWorkbook(sampleDataToWrite, "sample")
    val out = FileOutputStream(localPath)

    with(workbook) {
      write(out)
      out.close()
      dispose()
      close()
    }
    excelFile = File(localPath)
  }

  afterSpec {
    excelFile.delete()
  }

  should("write excel file") {
    excelFile.exists() shouldBe true
  }

  should("read excel file") {
    val sampleDataAfterRead = ExcelReader(localPath).readExcelFile<SampleDto>()
    sampleDataAfterRead.forEachIndexed { index, sampleDto ->

//      println("sampleDto: $sampleDto")
//      println("sampleDataAfterRead[index]: ${sampleDataAfterRead[index]}")

      sampleDto shouldBe sampleDataAfterRead[index]
    }
  }
})
