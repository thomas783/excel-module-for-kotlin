package shared

abstract class ExcelWriterCommonDto<T> {
  abstract fun createSampleData(size: Int): Collection<T>
}
