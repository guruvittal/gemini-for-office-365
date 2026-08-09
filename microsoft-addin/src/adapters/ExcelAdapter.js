/**
 * Host Adapter Implementation for Microsoft Excel (Excel.run)
 * 
 * @author Sathya AG, Principal Architect, Google
 */
export class ExcelAdapter {
  constructor() {
    this.name = "Excel";
  }

  // Read currently selected cell range text in Excel
  async getSelectedText() {
    let selectedText = "";
    try {
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load("values");
          await context.sync();

          if (range.values && Array.isArray(range.values)) {
            selectedText = range.values
              .map(row => row.join("\t"))
              .filter(line => line.trim().length > 0)
              .join("\n");
          }
        });
      }
    } catch (err) {
      console.warn("Excel selection read error:", err);
    }
    return selectedText;
  }

  // Read current sheet text context
  async getFullDocumentText() {
    return this.getSelectedText();
  }

  // Insert AI table/grid data into Excel active worksheet cells & overlay chart images
  async insertContent(htmlContent) {
    try {
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = context.workbook.getSelectedRange();
          range.load("rowIndex, columnIndex");
          await context.sync();

          const tempDiv = document.createElement("div");
          tempDiv.innerHTML = htmlContent;

          // Parse HTML Table into Native Excel Cell Matrices
          const table = tempDiv.querySelector("table");
          if (table) {
            const rows = Array.from(table.querySelectorAll("tr"));
            const matrix = rows.map(tr => {
              const cells = Array.from(tr.querySelectorAll("th, td"));
              return cells.map(td => td.innerText ? td.innerText.trim() : "");
            });

            if (matrix.length > 0) {
              const rowCount = matrix.length;
              const colCount = Math.max(...matrix.map(r => r.length));
              const targetRange = sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, rowCount, colCount);
              
              targetRange.values = matrix;
              targetRange.format.font.name = "Segoe UI";
              targetRange.format.font.size = 11;
              targetRange.format.autofitColumns();
              await context.sync();
            }
          } else {
            // Text fallback insertion into active cell
            const plainText = tempDiv.innerText || tempDiv.textContent || "";
            if (plainText.trim()) {
              range.values = [[plainText.trim()]];
              range.format.font.name = "Segoe UI";
              await context.sync();
            }
          }
        });
      }
    } catch (err) {
      console.error("Excel insertion error:", err);
    }
  }

  // Scan in-cell @gemini commands for Excel
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    if (callbacks.onStatus) callbacks.onStatus("Excel Adapter Ready (@gemini in cell)");
  }
}
