import org.gradle.api.tasks.testing.logging.TestExceptionFormat
import org.gradle.api.tasks.testing.logging.TestLogEvent

plugins {
  kotlin("jvm") version "1.9.24"
}

group = "com.excelkotlin"
version = "1.0-SNAPSHOT"

repositories {
  mavenCentral()
}

dependencies {
  // excel
  implementation("org.apache.poi:poi:4.1.1")
  implementation("org.apache.poi:poi-ooxml:4.1.1")

  // reflect
  implementation(kotlin("reflect"))

  // type parser
  implementation("com.github.drapostolos:type-parser:0.8.1")

  // commons
  implementation("org.apache.commons:commons-lang3:3.12.0")

  // validation
  implementation("javax.validation:validation-api:2.0.1.Final")

  // test
  testImplementation("io.kotest:kotest-runner-junit5:5.8.0")
  testImplementation("org.assertj:assertj-core:3.25.3")
  testImplementation("org.junit.jupiter:junit-jupiter:5.10.2")
}

tasks {
  test {
    useJUnitPlatform()
    testLogging {
      events = setOf(TestLogEvent.FAILED, TestLogEvent.SKIPPED, TestLogEvent.PASSED)
      showStandardStreams = true
      exceptionFormat = TestExceptionFormat.FULL
    }
  }
}

kotlin {
  jvmToolchain(21)
}
