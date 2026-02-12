#!/usr/bin/env node

const fs = require("node:fs")
const fsp = require("node:fs/promises")
const path = require("node:path")
const os = require("node:os")
const https = require("node:https")

const pkg = require("../package.json")

async function main() {
  const triple = supportedTriple(process.platform, process.arch)
  if (!triple) {
    throw new Error(`Unsupported platform: ${process.platform}/${process.arch}`)
  }

  const version = pkg.version
  const asset = `agent-spreadsheet-${triple.asset}`

  const base = process.env.AGENT_SPREADSHEET_DOWNLOAD_BASE_URL ||
    "https://github.com/PSU3D0/spreadsheet-mcp/releases/download"

  const url = `${base}/v${version}/${asset}`

  const vendorDir = path.join(__dirname, "..", "vendor")
  await fsp.mkdir(vendorDir, { recursive: true })

  const dest = path.join(vendorDir, triple.dest)
  await download(url, dest)

  if (process.platform !== "win32") {
    await fsp.chmod(dest, 0o755)
  }
}

function supportedTriple(platform, arch) {
  if (platform === "linux" && arch === "x64") {
    return { asset: "linux-x86_64", dest: "agent-spreadsheet" }
  }
  if (platform === "darwin" && arch === "x64") {
    return { asset: "macos-x86_64", dest: "agent-spreadsheet" }
  }
  if (platform === "darwin" && arch === "arm64") {
    return { asset: "macos-aarch64", dest: "agent-spreadsheet" }
  }
  if (platform === "win32" && arch === "x64") {
    return { asset: "windows-x86_64.exe", dest: "agent-spreadsheet.exe" }
  }
  return null
}

function download(url, dest) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest)
    const request = https.get(url, (response) => {
      if (response.statusCode && response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        file.close()
        fs.unlink(dest, () => {
          download(response.headers.location, dest).then(resolve, reject)
        })
        return
      }

      if (response.statusCode !== 200) {
        file.close()
        fs.unlink(dest, () => {
          reject(new Error(`Failed to download ${url} (status ${response.statusCode})`))
        })
        return
      }

      response.pipe(file)
      file.on("finish", () => file.close(resolve))
    })

    request.on("error", (error) => {
      file.close()
      fs.unlink(dest, () => reject(error))
    })
  })
}

main().catch((error) => {
  console.error(`[agent-spreadsheet] install failed: ${error.message}`)
  process.exit(1)
})
