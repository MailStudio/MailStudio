const { ipcRenderer } = require('electron')

const serviceArg = process.argv.find((value) => value.startsWith('--service-key='))
const serviceKey = serviceArg ? serviceArg.slice('--service-key='.length) : 'unknown'

function sendPageMeta() {
  ipcRenderer.send('service:page-meta', {
    serviceKey,
    title: document.title,
    href: window.location.href
  })
}

window.addEventListener('DOMContentLoaded', () => {
  sendPageMeta()

  const titleNode = document.querySelector('title')
  if (titleNode) {
    const observer = new MutationObserver(() => sendPageMeta())
    observer.observe(titleNode, { childList: true, subtree: true, characterData: true })
  }

  window.addEventListener('focus', sendPageMeta)
  window.addEventListener('hashchange', sendPageMeta)
})
