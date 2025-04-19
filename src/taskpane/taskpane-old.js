/* global document, Office, Word */

// async function insertTableFromText() {
//   try {
//     await Word.run(async (context) => {
//       const input = document.getElementById("textInput").value.trim();

//       if (!input) {
//         console.error("No input provided");
//         return;
//       }

//       const rows = input.split("\n").map((line) => line.split(",").map((cell) => cell.trim()));

//       if (rows.length === 0 || rows[0].length === 0) {
//         console.error("Input format is incorrect");
//         return;
//       }

//       const table = context.document.body.insertTable(
//         rows.length,
//         rows[0].length,
//         Word.InsertLocation.start,
//         rows
//       );

//       table.styleBuiltIn = Word.Style.gridTable5Dark_Accent1;

//       await context.sync();
//     });
//   } catch (error) {
//     console.error("Error inserting table:", error);
//   }
// }

async function insertTableFromText() {
  const errorElement = document.getElementById("error-message");
  errorElement.style.display = "none"; // Hide previous error

  try {
    await Word.run(async (context) => {
      const input = document.getElementById("textInput").value.trim();

      // if (!input) {
      //   errorElement.textContent = "Please enter some text.";
      //   errorElement.style.display = "block";
      //   return;
      // }

      const rows = input
        .split("\n")
        .map((line) => line.split(",").map((cell) => cell.trim()))
        .filter((row) => row.length > 0);

      const columnCount = rows[0]?.length || 0;
      const hasInvalidRows = rows.some((row) => row.length !== columnCount);

      if (rows.length === 0 || columnCount === 0 || hasInvalidRows) {
        errorElement.textContent =
          "Input format is incorrect. Each line should have the same number of comma-separated values.";
        errorElement.style.display = "block";
        return;
      }

      const table = context.document.body.insertTable(
        rows.length,
        columnCount,
        Word.InsertLocation.start,
        rows
      );

      table.styleBuiltIn = Word.Style.gridTable5Dark_Accent1;

      await context.sync();
    });
  } catch (error) {
    console.error("Error inserting table:", error);
    errorElement.textContent = "Something went wrong. Please try again.";
    errorElement.style.display = "block";
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Show main UI
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";
  }
});

// Export insert function to global scope so button can call it
window.insertTableFromText = insertTableFromText;
