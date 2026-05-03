Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insertBtn").onclick = insertText;
  }
});

function insertText() {
  alert("Button clicked!");
  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(
      "Hello from my first Word Add-in!",
      Word.InsertLocation.replace
    );
    await context.sync();
  });
}
