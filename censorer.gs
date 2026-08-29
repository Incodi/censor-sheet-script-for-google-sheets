function updatePublicSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const source = ss.getSheetByName("other youtubers");
  const censorSheet = ss.getSheetByName("censored");
  const publicSheet = ss.getSheetByName("youtuber phrases, public and clean (loads slowly)");

  if (!source) {
  throw new Error('Your source sheet doesnt exist');
  }

  if (!censorSheet) {
    throw new Error('Your censor list sheet doesnt exist.');
  }

  if (!publicSheet) {
    throw new Error('Your public sheet doesnt exist');
  }

  const censorList = censorSheet.getLastRow();

  const swearWords = censorList >= 2
    ? censorSheet
        .getRange(2, 1, censorList - 1, 1)
        .getValues()
        .flat()
        .map(String)
        .map(word => word.trim())
        .filter(Boolean)
    : [];

  const sourceRows = source.getLastRow();
  const sourceColumns = source.getLastColumn();

  const data = source
    .getRange(1, 1, sourceRows, sourceColumns)
    .getValues();
 
  let censorRegex = null;

  if (swearWords.length > 0) {
    const escaped = swearWords.map(word =>
      word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    );

    censorRegex = new RegExp(
      "\\b(" + escaped.join("|") + ")\\b",
      "gi"
    );
  }

  // Censoring ONLY column E
  const publicData = data.map((row, rowIndex) => {
    const newRow = [...row];

    if (rowIndex > 0 && censorRegex) {
      // Column E = index 4
      if (typeof newRow[4] === "string" && censorRegex.test(newRow[4])) {
        newRow[4] = "[Censored]";
      }
      censorRegex.lastIndex = 0;

      // Column F (if you want to censor column F too)

      // if (typeof newRow[5] === "string" && censorRegex.test(newRow[5])) {
      //   newRow[5] = "[Censored]";
      // }
      // censorRegex.lastIndex = 0;
    }

    return newRow;
  });

  // code for cenoring all columns
  //   const publicData = data.map((row, rowIndex) => {
  //   const newRow = [...row];

  //   if (rowIndex > 0 && censorRegex) {
  //     for (let colIndex = 0; colIndex < newRow.length; colIndex++) {
  //       const cellValue = newRow[colIndex];
        
  //       if (typeof cellValue === "string") {
  //         if (censorRegex.test(cellValue)) {
  //           newRow[colIndex] = "[Censored]";
  //         }
  //         censorRegex.lastIndex = 0;
  //       }
  //     }
  //   }

  //   return newRow;
  // });

  // Replace public sheet contents
  publicSheet.clearContents();

  publicSheet
    .getRange(1, 1, publicData.length, publicData[0].length)
    .setValues(publicData);
}