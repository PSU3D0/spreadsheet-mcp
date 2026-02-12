#!/usr/bin/env node

const { spawn } = require("node:child_process")
const fs = require("node:fs")
const path = require("node:path")

const isWin = process.platform === "win32"
const binaryName = isWin ? "agent-spreadsheet.exe" : "agent-spreadsheet"
const binPath = path.join(__dirname, "..", "vendor", binaryName)

if (!fs.existsSync(binPath)) {
  console.error(
    JSON.stringify({
      code: "BINARY_NOT_INSTALLED",
      message: "agent-spreadsheet binary not found; reinstall package to fetch platform binary",
      try_this: "npm i -g agent-spreadsheet"
    })
  )
  process.exit(1)
}

const child = spawn(binPath, process.argv.slice(2), {
  stdio: "inherit",
  windowsHide: true
})

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal)
    return
  }
  process.exit(code ?? 1)
})
