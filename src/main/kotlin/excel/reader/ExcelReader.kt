package excel.reader

import excel.reader.exception.ExcelReaderException
import org.apache.commons.collections4.ListUtils
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.util.*
import java.util.function.Function
import kotlin.reflect.KMutableProperty1

class ExcelReader {

  data class ExcelReaderItem(
    var errorFieldList: MutableList<ExcelReaderErrorField>,
    var workbook: Workbook,
  )

  data class ExcelHeaderValue<T>(
    val headerName: String,
    val headerIdx: Int,
    val field: KMutableProperty1<T, *>
  )

  companion object {
    lateinit var excelReaderItem: ExcelReaderItem

    fun initExcelReaderItem(path: String) {
      val excelFile = File(path).also {
        checkFileValidation(it)
      }
      val workbook: Workbook = runCatching {
        WorkbookFactory.create(excelFile)
      }.onFailure {
        throw ExcelReaderException(it.message.toString())
      }.getOrThrow()

      excelReaderItem = ExcelReaderItem(
        errorFieldList = mutableListOf(),
        workbook = workbook,
      )
    }

    @Throws(ExcelReaderException::class)
    private fun checkFileValidation(file: File) {
      val fileExtension = file.name.substring(file.name.lastIndexOf(".") + 1)

      if (fileExtension != "xlsx" && fileExtension != "xls")
        throw ExcelReaderException("Invalid file extension. Only .xlsx or .xls file is allowed.")
    }

    inline fun <reified T : Any> getObjectList(startRow: Int = 1, rowFunc: Function<Row, T>): List<T> {
      if (Objects.isNull(rowFunc)) throw ExcelReaderException("No row function to process.")

      val (errorFieldList, workbook) = excelReaderItem
      val sheet = workbook.getSheetAt(0)
      val rowCount = sheet.physicalNumberOfRows
      val objectList = (startRow until rowCount)
        .filter { rowIdx -> isRowAllBlank(sheet.getRow(rowIdx)) }
        .map { rowIdx -> rowFunc.apply(sheet.getRow(rowIdx)) }

      if (ListUtils.emptyIfNull(errorFieldList).isNotEmpty())
        throw ExcelReaderException("Something went wrong while reading the excel file. ${errorFieldList.joinToString("\n") { it.toString() }}")

      return objectList
    }

    fun isRowAllBlank(row: Row): Boolean {
      return row.cellIterator().asSequence().all { it.cellType == CellType.BLANK }
    }
  }
}
