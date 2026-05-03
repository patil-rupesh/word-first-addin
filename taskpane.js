console.log("✅ taskpane.js loaded");

Office.onReady((info) => {
  console.log("✅ Office.onReady fired", info);

  if (info.host === Office.HostType.Word) {
    const btn = document.getElementById("insertBtn");
    console.log("🔍 Button found:", btn);

    if (btn) {
      btn.onclick = insertText;
      console.log("✅ Click handler attached");
    }
  }
});

function insertText() {
  console.log("✅ Insert button clicked");

  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(
      "Hello from my first Word Add-in!",
      Word.InsertLocation.replace
    );
    await context.sync();
  });
}
