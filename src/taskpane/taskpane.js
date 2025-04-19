async function insertTableFromSelection() {
  const errorEl = document.getElementById("error-message");
  errorEl.style.display = "none";

  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("paragraphs");
      await context.sync();

      // Build a 2D array of [ [col1, col2,…], [ … ], … ]
      const rows = sel.paragraphs.items
        .map((p) => p.text.trim())
        .filter((line) => line.length > 0)
        .map((line) => line.split(",").map((c) => c.trim()));

      if (!rows.length) {
        throw new Error("Select comma‑separated lines first.");
      }
      const colCount = rows[0].length;
      if (rows.some((r) => r.length !== colCount)) {
        throw new Error("All rows must have the same number of columns.");
      }

      // 1) Insert the table *after* the selection
      const table = sel.insertTable(rows.length, colCount, Word.InsertLocation.after, rows);
      table.styleBuiltIn = Word.Style.gridTable5Dark_Accent1;

      await context.sync();

      // 2) Delete the original paragraphs to emulate "replace"
      for (const p of sel.paragraphs.items) {
        p.clear(); // or p.delete() in newer APIs
      }
      await context.sync();
    });
  } catch (err) {
    console.error("Error inserting table:", err);
    document.getElementById("error-message").textContent = err.message;
    document.getElementById("error-message").style.display = "block";
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";
  }
});

window.insertTableFromSelection = insertTableFromSelection;
