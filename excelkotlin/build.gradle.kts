dependencies {
  implementation("org.apache.poi:poi:4.1.1")
  implementation("org.apache.poi:poi-ooxml:4.1.1")

  // reflect
  implementation(kotlin("reflect"))

  // type parser
  implementation("com.github.drapostolos:type-parser:0.8.1")

  // commons
  implementation("org.apache.commons:commons-lang3:3.12.0")

  // validation
  implementation("org.valiktor:valiktor-core:0.12.0")

  // test
  testImplementation("io.kotest:kotest-runner-junit5:5.8.0")
  testImplementation("org.assertj:assertj-core:3.25.3")
  testImplementation("org.junit.jupiter:junit-jupiter:5.10.2")
}
