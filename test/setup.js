// jsdom has no ClipboardEvent; EditorView.pasteHTML/pasteText construct one.
if (typeof globalThis.ClipboardEvent === 'undefined') {
  globalThis.ClipboardEvent = class ClipboardEvent extends Event {
    constructor(type, init = {}) {
      super(type, init)
      this.clipboardData = init.clipboardData ?? null
    }
  }
}
