/**
 * Identify which application produced clipboard HTML.
 * @returns {'excel'|'google'|'libreoffice'|'numbers'|'generic'|null} null when there is no table
 */
export function detectSource(doc, html = '') {
  if (doc.querySelector('google-sheets-html-origin, [data-sheets-root]') || /google-sheets-html-origin/i.test(html)) {
    return 'google'
  }
  let progId = ''
  let generator = ''
  for (const meta of doc.querySelectorAll('meta')) {
    const name = (meta.getAttribute('name') || '').toLowerCase()
    if (name === 'progid') progId = meta.getAttribute('content') || ''
    if (name === 'generator') generator = meta.getAttribute('content') || ''
  }
  if (/excel\.sheet/i.test(progId) || /microsoft excel/i.test(generator)) return 'excel'
  // Excel Online / older builds may omit the meta but keep the namespace. Word also declares
  // office namespaces, so exclude it explicitly.
  if (/urn:schemas-microsoft-com:office:excel/i.test(html) && !/word\.document|microsoft word/i.test(progId + generator + html.slice(0, 2000))) {
    return 'excel'
  }
  if (/libreoffice|openoffice/i.test(generator)) return 'libreoffice'
  if (!doc.querySelector('table')) return null
  if (/cocoa html writer|numbers/i.test(generator)) return 'numbers'
  return 'generic'
}
