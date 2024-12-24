package shared

import io.kotest.core.annotation.Ignored
import io.kotest.core.spec.DslDrivenSpec
import io.kotest.core.spec.Spec
import java.io.File

@Ignored
class ExcelReaderBaseTests(
  private val path: String,
) : DslDrivenSpec() {
  lateinit var excelFile: File

  private val localPath: String
    get() = "src/test/resources/$path.xlsx"

  init {
    beforeSpec {
      excelFile = File(localPath)
    }

    afterSpec {
      excelFile.delete()
    }
  }

  companion object {
    fun Spec.setExcelReaderCommonSpec(
      path: String
    ): ExcelReaderBaseTests = ExcelReaderBaseTests(path)
  }
}
