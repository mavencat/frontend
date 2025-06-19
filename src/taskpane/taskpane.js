/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("highlight-range").onclick = highlightRange;
        document.getElementById("insert-text").onclick = insertText;
        document.getElementById("get-range-info").onclick = getRangeInfo;
        
        logMessage("Office.js is ready for Excel");
    } else {
        logMessage("This add-in is designed for Excel");
    }
});

/**
 * Highlights the selected range in yellow
 */
async function highlightRange() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "yellow";
            
            await context.sync();
            logMessage("Selected range highlighted in yellow");
        });
    } catch (error) {
        logError("Failed to highlight range", error);
    }
}

/**
 * Inserts sample text into the selected range
 */
async function insertText() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.values = [["Hello from Office.js!"]];
            
            await context.sync();
            logMessage("Sample text inserted");
        });
    } catch (error) {
        logError("Failed to insert text", error);
    }
}

/**
 * Gets information about the selected range
 */
async function getRangeInfo() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            
            // Load properties we want to read
            range.load("address");
            range.load("rowCount");
            range.load("columnCount");
            range.load("values");
            
            await context.sync();
            
            const info = `
Selected Range Information:
- Address: ${range.address}
- Rows: ${range.rowCount}
- Columns: ${range.columnCount}
- Values: ${JSON.stringify(range.values, null, 2)}`;
            
            logMessage(info);
        });
    } catch (error) {
        logError("Failed to get range info", error);
    }
}

/**
 * Logs a message to the output area
 * @param {string} message The message to log
 */
function logMessage(message) {
    const output = document.getElementById("output");
    const timestamp = new Date().toLocaleTimeString();
    output.textContent = `[${timestamp}] ${message}\n` + output.textContent;
}

/**
 * Logs an error message to the output area
 * @param {string} message The error message
 * @param {Error} error The error object
 */
function logError(message, error) {
    const errorMessage = `ERROR: ${message}\nDetails: ${error.message}`;
    logMessage(errorMessage);
    console.error(message, error);
}