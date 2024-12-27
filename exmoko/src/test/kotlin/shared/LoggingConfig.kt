package shared

import io.kotest.common.ExperimentalKotest
import io.kotest.core.config.AbstractProjectConfig
import io.kotest.core.config.LogLevel
import io.kotest.core.extensions.Extension
import io.kotest.core.test.TestCase
import io.kotest.engine.test.logging.LogEntry
import io.kotest.engine.test.logging.LogExtension

object LoggingConfig : AbstractProjectConfig() {
  override val logLevel = if (isDebugMode()) LogLevel.Debug else LogLevel.Info

  @OptIn(ExperimentalKotest::class)
  override fun extensions(): List<Extension> = super.extensions() + listOf(
    object : LogExtension {
      override suspend fun handleLogs(testCase: TestCase, logs: List<LogEntry>) {
        logs.forEach {
          println(it.level.name + " - " + it.message)
        }
      }
    }
  )

  private fun isDebugMode(): Boolean {
    return System.getenv("DEBUG_MODE") == "true" || System.getProperty("debug") == "true"
  }
}
