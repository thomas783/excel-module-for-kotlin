plugins {
  id("org.gradle.toolchains.foojay-resolver-convention") version "0.5.0"
}

rootProject.name = "excelkotlin-multi-module-app"

include("excelkotlin")
include("excelkotlin-sample")
