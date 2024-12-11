package shared

interface IExcelWriterCommonDto<T> {
  fun createSampleData(size: Int): Collection<T>
}
