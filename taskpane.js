function insertText() {
  Word.run(function (context) {
    const range = context.document.getSelection();
    range.insertText("Hello from my first Word Add-in!", Word.InsertLocation.replace);
    return context.sync();
  });
}