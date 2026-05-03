Office.onReady((info) => {
  alert("✅ Office is ready");

  if (info.host === Office.HostType.Word) {
    const btn = document.getElementById("insertBtn");
    if (btn) {
      btn.onclick = insertText;
      alert("✅ Button handler attached");
    } else {
      alert("❌ Button not found in DOM");
    }
  }
});

function insertText() {
  alert("✅ Button clicked!");

  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(
      "Hello from my first Word Add-in!",
      Word.InsertLocation.replace
    );
    await context.sync();
  });
}
