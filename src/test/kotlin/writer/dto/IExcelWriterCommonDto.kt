package writer.dto

interface IExcelWriterCommonDto<T> {
  fun createSampleData(size: Int): Collection<T>
}
