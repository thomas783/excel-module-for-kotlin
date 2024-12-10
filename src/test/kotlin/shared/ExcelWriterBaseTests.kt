package shared

import excel.writer.ExcelWriter
import io.kotest.core.spec.DslDrivenSpec
import io.kotest.core.spec.Spec
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileOutputStream

class ExcelWriterBaseTests(
  val sheetName: String = "Sample Excel File",
  val localPath: String = "src/test/resources/sample.xlsx",
) : DslDrivenSpec() {
  lateinit var createdFile: File
  lateinit var createdWorkbook: Workbook

  inline fun <reified T : Any> setCommonSpec(spec: Spec, sampleData: Collection<T>) {
    spec.apply {

      beforeSpec {
        val workbook = ExcelWriter.createWorkbook(sampleData, sheetName)
        val out = FileOutputStream(localPath)

        with(workbook) {
          write(out)
          out.close()
          dispose()
          close()
        }
        createdFile = File(localPath)
        createdWorkbook = createdFile.inputStream().use { WorkbookFactory.create(it) }
      }

      afterSpec {
        createdWorkbook.close()
        createdFile.delete()
      }
    }
  }
}
