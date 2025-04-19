/* global Office, Word */

async function insertTableFromSelection(event) {
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const paras = sel.paragraphs;
      paras.load("text");
      await context.sync();

      const rows = paras.items
        .map((p) => p.text.trim())
        .filter((line) => line.length > 0)
        .map((line) => line.split(",").map((cell) => cell.trim()));

      if (rows.length === 0) throw new Error("Select comma‑separated lines.");
      const colCount = rows[0].length;
      if (rows.some((r) => r.length !== colCount))
        throw new Error("All rows must have the same columns.");

      const table = sel.insertTable(rows.length, colCount, Word.InsertLocation.after, rows);
      table.styleBuiltIn = Word.Style.gridTable5Dark_Accent1;
      await context.sync();

      // Optionally delete original paragraphs to emulate replace:
      paras.items.forEach((p) => p.clear());
      await context.sync();
    });
  } catch (err) {
    console.error("insertTableFromSelection error:", err);
    // You can surface feedback to the user via a dialog or notification if you want
  } finally {
    // IMPORTANT: let Office know you’re done!
    event.completed();
  }
}

Office.onReady(() => {
  Office.actions.associate("insertTableFromSelection", insertTableFromSelection);
});
