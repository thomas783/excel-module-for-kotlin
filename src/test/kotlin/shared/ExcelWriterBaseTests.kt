package shared

import com.excelkotlin.writer.ExcelWriter
import io.kotest.core.annotation.Ignored
import io.kotest.core.spec.DslDrivenSpec
import io.kotest.core.spec.Spec
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileOutputStream
import kotlin.reflect.full.createInstance

@Ignored
class ExcelWriterBaseTests(
  val sampleDataSize: Int = 1000,
  val sheetName: String,
  private val path: String,
  val spec: Spec,
  initialize: ExcelWriterBaseTests.() -> Unit,
) : DslDrivenSpec() {
  lateinit var excelFile: File
  lateinit var workbook: Workbook
  val localPath: String
    get() = "src/test/resources/sample/$path.xlsx"

  init {
    initialize()
  }

  inline fun <reified T : IExcelWriterCommonDto<K>, reified K : Any> getSampleData(): Collection<K> {
    val instance = T::class.objectInstance ?: T::class.createInstance()
    return instance.createSampleData(sampleDataSize)
  }

  inline fun <reified T : IExcelWriterCommonDto<K>, reified K : Any> setCommonSpec() {
    val sampleData = getSampleData<T, K>()
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
        excelFile = File(localPath)
      }

      beforeTest {
        workbook = excelFile.inputStream().use { WorkbookFactory.create(it) }
      }

      afterTest {
        workbook.close()
      }

      afterSpec {
        excelFile.delete()
      }
    }
  }

  companion object {
    inline fun <reified T : IExcelWriterCommonDto<K>, reified K : Any> Spec.setExcelWriterCommonSpec(
      sampleDataSize: Int = 1000,
      sheetName: String = "Sample Excel File",
      path: String = "sample",
    ): ExcelWriterBaseTests {
      val baseTest = ExcelWriterBaseTests(
        sampleDataSize = sampleDataSize,
        sheetName = sheetName,
        path = path,
        spec = this,
      ) {
        this.setCommonSpec<T, K>()
      }

      return baseTest
    }
  }
}
