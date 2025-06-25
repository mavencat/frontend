/* global console, document, Excel, Office */

// Import CSS
import './taskpane.css';

// Chat state
const chatState = {
    messages: [],
    isLoading: false
};

// Initialize when Office.js is ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeChat();
        console.log("Office.js is ready for Excel - Chat interface initialized");
    } else {
        showError("This add-in is designed for Excel");
    }
});

/**
 * Initialize chat interface and workbook data
 */
async function initializeChat() {
    const messageInput = document.getElementById('message-input');
    const sendButton = document.getElementById('send-button');
    
    sendButton.addEventListener('click', handleSendMessage);
    messageInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSendMessage();
        }
    });
    
    try {
        setLoadingState(true, 'Initializing workbook data...');
        
        // Get workbook name and worksheet information
        const workbookName = await getSimpleWorkbookName();
        const worksheetData = await getWorksheetNames();
        
        // Immediately extract cell data after getting worksheet names
        console.log('Starting cell data extraction...');
        setLoadingState(true, 'Extracting cell data from workbook...');
        
        const cellData = await extractAllWorksheetData();
        console.log(`Extracted ${cellData.length} cells from workbook`);
        
        const workbookData = {
            workbookName: workbookName,
            totalWorksheets: worksheetData.length,
            sheets: worksheetData
        };
        
        console.log('Workbook data prepared:', workbookData);
        const result = await sendWorkbookData(workbookData);
        currentFileId = result.file_id;
        console.log('Workbook initialization complete:', result);
        
        // Transmit the cell data we extracted
        if (currentFileId && cellData.length > 0) {
            console.log('Transmitting cell data to backend...');
            setLoadingState(true, 'Transmitting cell data to backend...');
            
            try {
                const cellDataResult = await transmitAllCellData(currentFileId, cellData, showProgress);
                console.log('Cell data transmission completed:', cellDataResult);
                
                if (cellDataResult.success) {
                    addMessage(`🎉 Workbook ready! Successfully extracted and transmitted ${cellDataResult.totalCells} cells from your Excel model. You can now chat about your data.`, 'ai');
                } else {
                    addMessage(`⚠️ Workbook partially ready. ${cellDataResult.successfulBatches}/${cellDataResult.totalBatches} batches transmitted successfully. Some data may be missing from analysis.`, 'ai');
                }
                
            } catch (cellError) {
                console.error('Cell data transmission failed:', cellError);
                addMessage('❌ Workbook initialized, but cell data transmission failed. You can still chat, but I won\'t have access to your cell data.', 'ai');
            }
        } else if (cellData.length === 0) {
            addMessage('📊 Workbook initialized. No cell data found to extract. Ready to chat!', 'ai');
        }
        
    } catch (error) {
        console.error('Workbook initialization failed:', error);
        showError('Failed to initialize workbook data');
    } finally {
        setLoadingState(false);
        messageInput.focus();
    }
}

/**
 * Handle sending a message
 */
async function handleSendMessage() {
    const messageInput = document.getElementById('message-input');
    const message = messageInput.value.trim();
    
    if (!message || chatState.isLoading) return;
    
    // Add user message
    addMessage(message, 'user');
    
    // Clear input
    messageInput.value = '';
    
    // Set loading state
    setLoadingState(true);
    
    try {
        // Get AI response (placeholder for now)
        const response = await getAIResponse(message);
        addMessage(response, 'ai');
    } catch (error) {
        addMessage('Sorry, I encountered an error. Please try again.', 'ai');
        console.error('Error getting AI response:', error);
    } finally {
        setLoadingState(false);
        messageInput.focus();
    }
}

/**
 * Add a message to the chat
 * @param {string} text - Message text
 * @param {string} sender - 'user' or 'ai'
 */
function addMessage(text, sender) {
    const message = {
        id: Date.now(),
        text: text,
        sender: sender,
        timestamp: new Date()
    };
    
    chatState.messages.push(message);
    renderMessage(message);
    scrollToBottom();
}

/**
 * Render a message in the chat
 * @param {Object} message - Message object
 */
function renderMessage(message) {
    const messagesContainer = document.getElementById('messages');
    
    const messageElement = document.createElement('div');
    messageElement.className = `message ${message.sender}-message`;
    messageElement.setAttribute('data-message-id', message.id);
    
    const contentElement = document.createElement('div');
    contentElement.className = 'message-content';
    
    const textElement = document.createElement('p');
    textElement.textContent = message.text;
    
    const timeElement = document.createElement('div');
    timeElement.className = 'message-time';
    timeElement.textContent = formatTime(message.timestamp);
    
    contentElement.appendChild(textElement);
    messageElement.appendChild(contentElement);
    messageElement.appendChild(timeElement);
    
    messagesContainer.appendChild(messageElement);
}

/**
 * Set loading state with custom message
 * @param {boolean} isLoading - Loading state
 * @param {string} message - Custom loading message
 */
function setLoadingState(isLoading, message = 'AI is thinking...') {
    chatState.isLoading = isLoading;
    const sendButton = document.getElementById('send-button');
    const messageInput = document.getElementById('message-input');
    const status = document.getElementById('status');
    
    sendButton.disabled = isLoading;
    messageInput.disabled = isLoading;
    
    if (isLoading) {
        status.className = 'status-typing';
        status.innerHTML = `<span class="typing-indicator">${message}</span>`;
    } else {
        status.className = 'status-hidden';
        status.textContent = '';
    }
}

/**
 * Show detailed progress for cell data extraction and transmission
 * @param {Object} progress - Progress information
 */
function showProgress(progress) {
    const status = document.getElementById('status');
    
    if (progress.stage === 'extraction_start') {
        status.className = 'status-typing';
        status.innerHTML = `<span class="typing-indicator">${progress.message}</span>`;
        
    } else if (progress.stage === 'extraction_complete') {
        status.className = 'status-typing';
        status.innerHTML = `<span class="typing-indicator">Extracted ${progress.totalCells} cells. Starting transmission...</span>`;
        
    } else if (progress.stage === 'preparation') {
        status.className = 'status-typing';
        status.innerHTML = `<span class="typing-indicator">Preparing ${progress.totalBatches} batches for ${progress.totalCells} cells...</span>`;
        
    } else if (progress.stage === 'transmission') {
        const percentage = Math.round((progress.completedBatches / progress.totalBatches) * 100);
        status.className = 'status-typing';
        status.innerHTML = `<span class="typing-indicator">Sending batch ${progress.completedBatches + 1}/${progress.totalBatches} (${percentage}%) - Sheet: ${progress.currentSheet}</span>`;
        
    } else if (progress.stage === 'complete') {
        const summary = progress.summary;
        if (summary.success) {
            status.className = 'status-success';
            status.innerHTML = `<span>✓ Successfully transmitted ${summary.totalCells} cells in ${summary.totalBatches} batches</span>`;
        } else {
            status.className = 'status-warning';
            status.innerHTML = `<span>⚠ Partial success: ${summary.successfulBatches}/${summary.totalBatches} batches sent. ${summary.failedBatches} failed.</span>`;
        }
        
        // Hide status after a few seconds
        setTimeout(() => {
            status.className = 'status-hidden';
            status.textContent = '';
        }, 5000);
        
    } else if (progress.stage === 'error') {
        status.className = 'status-error';
        status.textContent = `Error: ${progress.error}`;
        
        // Hide error after a longer time
        setTimeout(() => {
            status.className = 'status-hidden';
            status.textContent = '';
        }, 10000);
    }
}

/**
 * Test cell data extraction and transmission
 * This function can be called manually for testing
 */
async function testCellDataExtraction() {
    try {
        if (!currentFileId) {
            showError('No file ID available. Please initialize workbook first.');
            return;
        }
        
        console.log('Starting manual cell data extraction test...');
        
        // Extract and transmit cell data with progress tracking
        const result = await extractAndTransmitWorkbookData(currentFileId, showProgress);
        
        console.log('Manual cell data extraction test completed:', result);
        
        // Update extraction status
        cellDataExtracted = result.success;
        
        if (result.success) {
            addMessage(`🔄 Manual extraction completed successfully! Transmitted ${result.totalCells} cells.`, 'ai');
        } else {
            addMessage(`🔄 Manual extraction completed with issues. ${result.successfulBatches}/${result.totalBatches} batches successful.`, 'ai');
        }
        
    } catch (error) {
        console.error('Manual cell data extraction test failed:', error);
        showError('Cell data extraction failed: ' + error.message);
    }
}

/**
 * Force re-extraction of cell data (ignores cellDataExtracted flag)
 */
async function forceReExtractCellData() {
    try {
        if (!currentFileId) {
            showError('No file ID available. Please initialize workbook first.');
            return;
        }
        
        console.log('Starting forced cell data re-extraction...');
        cellDataExtracted = false; // Reset flag to force re-extraction
        
        // Extract and transmit cell data with progress tracking
        const result = await extractAndTransmitWorkbookData(currentFileId, showProgress);
        
        console.log('Forced cell data re-extraction completed:', result);
        
        // Update extraction status
        cellDataExtracted = result.success;
        
        if (result.success) {
            addMessage(`🔄 Re-extraction completed! Updated with ${result.totalCells} cells from your Excel model.`, 'ai');
        } else {
            addMessage(`🔄 Re-extraction partially completed. ${result.successfulBatches}/${result.totalBatches} batches successful.`, 'ai');
        }
        
    } catch (error) {
        console.error('Forced cell data re-extraction failed:', error);
        showError('Forced re-extraction failed: ' + error.message);
        cellDataExtracted = false; // Reset on failure
    }
}

// Make functions globally available for debugging
window.testCellDataExtraction = testCellDataExtraction;
window.forceReExtractCellData = forceReExtractCellData;

/**
 * Show error message
 * @param {string} message - Error message
 */
function showError(message) {
    const status = document.getElementById('status');
    status.className = 'status-error';
    status.textContent = message;
}

/**
 * Scroll chat to bottom
 */
function scrollToBottom() {
    const messagesContainer = document.getElementById('messages-container');
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}

/**
 * Format timestamp for display
 * @param {Date} date - Date object
 * @returns {string} Formatted time
 */
function formatTime(date) {
    const now = new Date();
    const diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000);
    
    if (diffMins < 1) return 'Just now';
    if (diffMins < 60) return `${diffMins}m ago`;
    
    return date.toLocaleTimeString([], { 
        hour: '2-digit', 
        minute: '2-digit' 
    });
}

// API configuration
const BASE_API_URL = 'https://backend-962119591036.europe-west1.run.app';
const CONFIG_URL = `${BASE_API_URL}/config`;
const INITIALIZE_URL = `${BASE_API_URL}/initialize`;
const CHAT_URL = `${BASE_API_URL}/chat`;
const STORE_CELL_DATA_URL = `${BASE_API_URL}/store-cell-data`;

// Configuration and state
let workbookConfig = null;
let currentFileId = null;
let cellDataExtracted = false;

/**
 * Fetch workbook processing configuration from backend
 */
async function fetchWorkbookConfig() {
    try {
        const response = await fetch(CONFIG_URL);
        if (!response.ok) {
            throw new Error(`Config fetch failed: ${response.status}`);
        }
        workbookConfig = await response.json();
        console.log('Workbook config loaded:', workbookConfig);
        return workbookConfig;
    } catch (error) {
        console.error('Failed to fetch workbook config:', error);
        workbookConfig = {};
        return workbookConfig;
    }
}

/**
 * Get simple workbook name without complex data extraction
 */
async function getSimpleWorkbookName() {
    return await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load('name');
        await context.sync();
        
        return workbook.name || 'Untitled Workbook';
    });
}

/**
 * Get worksheet names safely without complex data extraction
 */
async function getWorksheetNames() {
    try {
        return await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load(['name']);
            await context.sync();
            
            return worksheets.items.map(sheet => ({ sheet_name: sheet.name }));
        });
    } catch (error) {
        console.warn('Could not get worksheet names:', error);
        return [];
    }
}


/**
 * Send workbook data to backend
 */
async function sendWorkbookData(workbookData) {
    try {
        const response = await fetch(INITIALIZE_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify(workbookData)
        });
        
        if (!response.ok) {
            throw new Error(`Workbook initialization failed: ${response.status}`);
        }
        
        const result = await response.json();
        console.log('Workbook data sent successfully:', result);
        return result;
        
    } catch (error) {
        console.error('Failed to send workbook data:', error);
        throw error;
    }
}

/**
 * Get AI response from backend API
 * @param {string} userMessage - User's message
 * @returns {Promise<string>} AI response
 */
async function getAIResponse(userMessage) {
    
    try {
        // Get Excel context for AI
        const excelContext = await getExcelContext();
        
        // Prepare request payload
        const payload = {
            message: userMessage,
            context: excelContext,
            timestamp: new Date().toISOString(),
            file_id: currentFileId
        };
        
        console.log('Sending request to:', CHAT_URL);
        console.log('Payload:', payload);
        
        const response = await fetch(CHAT_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify(payload)
        });
        
        // Handle response
        if (!response.ok) {
            let errorMessage = `Server error (${response.status}): `;
            
            try {
                const errorData = await response.json();
                errorMessage += errorData.detail || 'Unknown error occurred';
            } catch (e) {
                errorMessage += `HTTP ${response.status} ${response.statusText}`;
            }
            
            throw new Error(errorMessage);
        }
        
        const data = await response.json();
        console.log('Response received:', data);
        
        // Validate response format
        if (!data.response) {
            throw new Error('Invalid response format: missing response field');
        }
        
        return data.response;
        
    } catch (error) {
        console.error('API Error:', error);
        throw new Error('Sorry, I encountered an error. Please try again.');
    }
}


/**
 * Get Excel context information (for future AI integration)
 * @returns {Promise<Object>} Excel context data
 */
async function getExcelContext() {
    try {
        return await Excel.run(async (context) => {
            const selectedRange = context.workbook.getSelectedRange();
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            
            selectedRange.load(['address']);
            worksheet.load(['name']);
            
            await context.sync();
            
            return {
                selectedRange: selectedRange.address,
                worksheet: worksheet.name
            };
        });
    } catch (error) {
        console.warn('Could not get Excel context:', error);
        return null;
    }
}

/**
 * Extract cell data from a worksheet's used range
 * @param {string} worksheetName - Name of the worksheet
 * @returns {Promise<Array>} Array of cell data objects
 */
async function getUsedRangeData(worksheetName) {
    try {
        return await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(worksheetName);
            const usedRange = worksheet.getUsedRange();
            
            if (!usedRange) {
                console.log(`No used range found in worksheet: ${worksheetName}`);
                return [];
            }
            
            usedRange.load(['values', 'formulas', 'text', 'address', 'rowCount', 'columnCount']);
            await context.sync();
            
            console.log(`Processing ${usedRange.rowCount}x${usedRange.columnCount} range in ${worksheetName}`);
            
            return convertRangeToCellData(usedRange, worksheetName);
        });
    } catch (error) {
        console.error(`Error extracting data from worksheet ${worksheetName}:`, error);
        return [];
    }
}

/**
 * Convert Excel range data to structured cell data format
 * @param {Object} range - Excel range object with loaded data
 * @param {string} sheetName - Name of the sheet
 * @returns {Array} Array of CellData objects
 */
function convertRangeToCellData(range, sheetName) {
    const cellDataArray = [];
    const values = range.values;
    const formulas = range.formulas;
    const text = range.text;
    const rowCount = range.rowCount;
    const columnCount = range.columnCount;
    
    for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < columnCount; col++) {
            const cellValue = values[row][col];
            const cellFormula = formulas[row][col];
            const cellText = text[row][col];
            
            // Skip empty cells to reduce data size
            if (cellValue === null && cellFormula === null && !cellText) {
                continue;
            }
            
            const cellData = {
                sheet_name: sheetName,
                column: getColumnLetter(col),
                row: row + 1, // Excel rows are 1-indexed
                formula: cellFormula !== cellValue ? cellFormula : null,
                value: cellValue,
                display_text: cellText
            };
            
            cellDataArray.push(cellData);
        }
    }
    
    console.log(`Extracted ${cellDataArray.length} non-empty cells from ${sheetName}`);
    return cellDataArray;
}

/**
 * Convert column index to Excel column letter
 * @param {number} columnIndex - 0-based column index
 * @returns {string} Column letter (A, B, C, ..., AA, AB, etc.)
 */
function getColumnLetter(columnIndex) {
    let columnLetter = '';
    while (columnIndex >= 0) {
        columnLetter = String.fromCharCode(65 + (columnIndex % 26)) + columnLetter;
        columnIndex = Math.floor(columnIndex / 26) - 1;
    }
    return columnLetter;
}

/**
 * Extract cell data from all worksheets in the workbook
 * @returns {Promise<Array>} Array of all cell data from all sheets
 */
async function extractAllWorksheetData() {
    try {
        const worksheetNames = await getWorksheetNames();
        console.log(`Found ${worksheetNames.length} worksheets to process`);
        
        let allCellData = [];
        
        for (const sheet of worksheetNames) {
            console.log(`Processing worksheet: ${sheet.sheet_name}`);
            const sheetCellData = await getUsedRangeData(sheet.sheet_name);
            allCellData = allCellData.concat(sheetCellData);
        }
        
        console.log(`Total cells extracted: ${allCellData.length}`);
        return allCellData;
        
    } catch (error) {
        console.error('Error extracting worksheet data:', error);
        throw error;
    }
}

/**
 * Batch cell data into manageable chunks for transmission
 * @param {Array} cellData - Array of cell data objects
 * @param {number} batchSize - Number of cells per batch
 * @returns {Array} Array of batched cell data
 */
function batchCellData(cellData, batchSize = 1500) {
    const batches = [];
    const cellsBySheet = {};
    
    // Group cells by sheet
    cellData.forEach(cell => {
        if (!cellsBySheet[cell.sheet_name]) {
            cellsBySheet[cell.sheet_name] = [];
        }
        cellsBySheet[cell.sheet_name].push(cell);
    });
    
    // Create batches for each sheet
    Object.keys(cellsBySheet).forEach(sheetName => {
        const sheetCells = cellsBySheet[sheetName];
        const totalCells = sheetCells.length;
        
        for (let i = 0; i < sheetCells.length; i += batchSize) {
            const batchCells = sheetCells.slice(i, i + batchSize);
            const batchNumber = Math.floor(i / batchSize) + 1;
            
            batches.push({
                sheet_name: sheetName,
                cells: batchCells,
                total_cells: totalCells,
                batch_number: batchNumber
            });
        }
    });
    
    console.log(`Created ${batches.length} batches from ${cellData.length} cells`);
    return batches;
}

/**
 * Send a batch of cell data to the backend
 * @param {string} fileId - File ID for the workbook
 * @param {Array} batches - Array of batched cell data
 * @param {boolean} isFinalBatch - Whether this is the final batch
 * @returns {Promise<Object>} Response from server
 */
async function sendCellDataBatch(fileId, batches, isFinalBatch = false) {
    try {
        const payload = {
            file_id: fileId,
            batches: batches,
            is_final_batch: isFinalBatch
        };
        
        console.log(`Sending batch with ${batches.length} batches to backend`);
        console.log(`Total cells in this transmission: ${batches.reduce((sum, batch) => sum + batch.cells.length, 0)}`);
        
        const response = await fetch(STORE_CELL_DATA_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify(payload)
        });
        
        if (!response.ok) {
            let errorMessage = `Server error (${response.status}): `;
            try {
                const errorData = await response.json();
                errorMessage += errorData.detail || 'Unknown error occurred';
            } catch (e) {
                errorMessage += `HTTP ${response.status} ${response.statusText}`;
            }
            throw new Error(errorMessage);
        }
        
        const result = await response.json();
        console.log('Batch sent successfully:', result);
        return result;
        
    } catch (error) {
        console.error('Error sending cell data batch:', error);
        throw error;
    }
}

/**
 * Send all cell data to backend in managed batches
 * @param {string} fileId - File ID for the workbook
 * @param {Array} cellData - Array of all cell data
 * @param {function} progressCallback - Optional callback for progress updates
 * @returns {Promise<Object>} Summary of transmission results
 */
async function transmitAllCellData(fileId, cellData, progressCallback = null) {
    try {
        console.log(`Starting transmission of ${cellData.length} cells for file ${fileId}`);
        
        // Create batches
        const batches = batchCellData(cellData, 1500);
        const totalBatches = batches.length;
        
        if (progressCallback) {
            progressCallback({
                stage: 'preparation',
                totalBatches: totalBatches,
                completedBatches: 0,
                currentSheet: null,
                totalCells: cellData.length
            });
        }
        
        let successfulBatches = 0;
        let failedBatches = 0;
        const errors = [];
        
        // Send batches sequentially to avoid overwhelming the server
        for (let i = 0; i < batches.length; i++) {
            const batch = batches[i];
            const isFinalBatch = (i === batches.length - 1);
            
            try {
                console.log(`Sending batch ${i + 1}/${totalBatches} for sheet: ${batch.sheet_name}`);
                
                if (progressCallback) {
                    progressCallback({
                        stage: 'transmission',
                        totalBatches: totalBatches,
                        completedBatches: i,
                        currentSheet: batch.sheet_name,
                        batchNumber: batch.batch_number,
                        cellsInBatch: batch.cells.length
                    });
                }
                
                // Send single batch wrapped in array
                await sendCellDataBatch(fileId, [batch], isFinalBatch);
                successfulBatches++;
                
                // Small delay between batches to prevent overwhelming the server
                if (!isFinalBatch) {
                    await new Promise(resolve => setTimeout(resolve, 100));
                }
                
            } catch (error) {
                console.error(`Failed to send batch ${i + 1}:`, error);
                failedBatches++;
                errors.push({
                    batchNumber: i + 1,
                    sheetName: batch.sheet_name,
                    error: error.message
                });
                
                // For now, continue with other batches even if one fails
                // In production, you might want to implement retry logic
            }
        }
        
        const summary = {
            totalCells: cellData.length,
            totalBatches: totalBatches,
            successfulBatches: successfulBatches,
            failedBatches: failedBatches,
            errors: errors,
            success: failedBatches === 0
        };
        
        if (progressCallback) {
            progressCallback({
                stage: 'complete',
                summary: summary
            });
        }
        
        console.log('Cell data transmission complete:', summary);
        return summary;
        
    } catch (error) {
        console.error('Error in cell data transmission:', error);
        throw error;
    }
}

/**
 * Extract and transmit all workbook cell data
 * @param {string} fileId - File ID for the workbook
 * @param {function} progressCallback - Optional callback for progress updates
 * @returns {Promise<Object>} Transmission summary
 */
async function extractAndTransmitWorkbookData(fileId, progressCallback = null) {
    try {
        if (progressCallback) {
            progressCallback({
                stage: 'extraction_start',
                message: 'Starting cell data extraction...'
            });
        }
        
        // Extract all cell data
        const cellData = await extractAllWorksheetData();
        
        if (cellData.length === 0) {
            console.log('No cell data found to transmit');
            return { success: true, totalCells: 0, message: 'No data to transmit' };
        }
        
        if (progressCallback) {
            progressCallback({
                stage: 'extraction_complete',
                totalCells: cellData.length,
                message: `Extracted ${cellData.length} cells, starting transmission...`
            });
        }
        
        // Transmit the data
        return await transmitAllCellData(fileId, cellData, progressCallback);
        
    } catch (error) {
        console.error('Error in extract and transmit:', error);
        if (progressCallback) {
            progressCallback({
                stage: 'error',
                error: error.message
            });
        }
        throw error;
    }
}

