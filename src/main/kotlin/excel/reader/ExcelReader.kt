package excel.reader

import excel.reader.annotation.ExcelReaderHeader
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
import kotlin.reflect.full.findAnnotation
import kotlin.reflect.full.memberProperties

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

    inline fun <reified T : Any> getHeader(rowNum: Int = 0): MutableMap<String, ExcelHeaderValue<T>> {
      val memberProperties = T::class.memberProperties.toList()
      val headers = excelReaderItem.workbook.getSheetAt(0).getRow(rowNum)
      val essentialHeaders = T::class.findAnnotation<ExcelReaderHeader>()?.essentialFields
      val readHeaders: MutableMap<String, ExcelHeaderValue<T>> =
        (0 until headers.physicalNumberOfCells).mapNotNull { cellIdx ->
          val headerName = headers.getCell(cellIdx).stringCellValue
          val field = memberProperties.firstOrNull { it.name == headerName } as? KMutableProperty1<T, *>?
          if (field != null) ExcelHeaderValue(headerName, cellIdx, field)
          else null
        }.associateBy { it.headerName }.toMutableMap()

      if (essentialHeaders != null) validateEssentialHeaders(essentialHeaders, readHeaders, rowNum)

      return readHeaders
    }

    @Throws(ExcelReaderException::class)
    fun <T : Any> validateEssentialHeaders(
      essentialHeaders: Array<String>,
      readHeaders: Map<String, ExcelHeaderValue<T>>,
      rowNum: Int
    ) {
      val error: ExcelReaderFieldError = ExcelReaderFieldError.HEADER_MISSING
      essentialHeaders.forEach { essentialHeader ->
        if (!readHeaders.keys.contains(essentialHeader)) excelReaderItem.errorFieldList.add(
          ExcelReaderErrorField(
            type = error.name,
            row = rowNum + 1,
            field = essentialHeader,
            fieldHeader = null,
            inputData = null,
            message = error.message,
            exceptionMessage = "$essentialHeader header is missing."
          )
        )
      }
      if (excelReaderItem.errorFieldList.isNotEmpty())
        throw ExcelReaderException("Essential headers are missing. ${excelReaderItem.errorFieldList.joinToString("\n") { it.toString() }}")
    }
  }
}
