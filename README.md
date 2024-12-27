![exmoko](docs/kotlin-logo.png)

> Exmoko is an Excel module for kotlin that allows you to read and write Excel files in kotlin.
> It is a wrapper around Apache POI library.

## Build

```shell
./gradlew build
```

## Getting Started

### To write an Excel file

```kotlin
import java.io.FileOutputStream

data class PersonToWrite(
  val name: String,
  val age: Int,
  val email: String,
  val phone: String
)

val path = "your/own/directory/persons.xlsx"
val sampleData = listOf(
  PersonToWrite("John Doe", 30, "johndoe@email.com", "1234567890"),
  PersonToWrite("John Snow", 25, "johnsnow@email.com", "828292929")
)
val outputStream = FileOutputStream(path)
val workbook = ExcelWriter.createWorkbook(sampleData, "sheet1").apply {
  write(outputStream)
  outputStream.close()
  dispose()
  close()
}
```

### To read an Excel file

```kotlin
data class PersonToRead(
  var name: String? = null,
  var age: Int? = null,
  var email: String? = null,
  var phone: String? = null
)

val path = "your/own/directory/persons.xlsx"
val persons = ExcelReader(path).readExcelFile<PersonToRead>()
```

See [sample](exmoko-sample) for more examples.

See [test-reader](exmoko/src/test/kotlin/reader) for better understanding of how to read Excel files.

See [test-writer](exmoko/src/test/kotlin/writer) for better understanding of how to write Excel files.
