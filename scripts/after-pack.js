const fs = require('fs')
const path = require('path')
const { execFileSync } = require('child_process')

function setPlistValue(plistPath, key, type, value) {
  execFileSync('plutil', ['-replace', key, `-${type}`, value, plistPath])
}

exports.default = async function afterPack(context) {
  if (context.electronPlatformName !== 'darwin') {
    return
  }

  const appName = context.packager.appInfo.productFilename
  const appBundlePath = path.join(context.appOutDir, `${appName}.app`)
  const appContentsPath = path.join(appBundlePath, 'Contents')
  const appIconPath = path.join(appContentsPath, 'Resources', 'icon.icns')
  const frameworksPath = path.join(appContentsPath, 'Frameworks')

  const helperNames = [
    'Outlook Orbit Helper',
    'Outlook Orbit Helper (GPU)',
    'Outlook Orbit Helper (Plugin)',
    'Outlook Orbit Helper (Renderer)'
  ]

  for (const helperName of helperNames) {
    const helperContentsPath = path.join(frameworksPath, `${helperName}.app`, 'Contents')
    const plistPath = path.join(helperContentsPath, 'Info.plist')
    const resourcesPath = path.join(helperContentsPath, 'Resources')
    const helperIconPath = path.join(resourcesPath, 'icon.icns')

    if (!fs.existsSync(plistPath)) {
      continue
    }

    fs.mkdirSync(resourcesPath, { recursive: true })
    fs.copyFileSync(appIconPath, helperIconPath)

    setPlistValue(plistPath, 'CFBundleName', 'string', helperName)
    setPlistValue(plistPath, 'CFBundleDisplayName', 'string', helperName)
    setPlistValue(plistPath, 'CFBundleIconFile', 'string', 'icon.icns')
  }

  // electron-builder is configured with `identity: null`, so it skips signing
  // and the bundle keeps the prebuilt Electron's stale signature (which no
  // longer matches after we edited Info.plists/icons above). Apply a fresh
  // ad-hoc signature to the whole bundle so it launches cleanly on Apple
  // Silicon (after the user clears Gatekeeper quarantine). Re-sign inner code
  // first, then the outer app, so each seal is consistent.
  execFileSync('codesign', ['--remove-signature', appBundlePath], { stdio: 'inherit' })
  execFileSync(
    'codesign',
    ['--force', '--deep', '--sign', '-', '--identifier', 'com.openai.outlookorbit', appBundlePath],
    { stdio: 'inherit' }
  )
  execFileSync('codesign', ['--verify', '--deep', '--strict', appBundlePath], { stdio: 'inherit' })
}
