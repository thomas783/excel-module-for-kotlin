plugins {
  id("org.gradle.toolchains.foojay-resolver-convention") version "0.5.0"
}

rootProject.name = "exmoko-multi-module-app"

include("exmoko")
include("exmoko-sample")
