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
        const workbookData = await getFullWorkbookData();
        const result = await sendWorkbookData(workbookData);
        currentWorkbookName = workbookData.workbookName;
        console.log('Workbook initialization complete:', result);
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

// Configuration and state
let workbookConfig = null;
let currentWorkbookName = null;

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
        workbookConfig = {
            MAX_ROWS_PER_BATCH: 250,
            MAX_TOTAL_ROWS: 2000,
            MAX_TOTAL_COLS: 104
        };
        return workbookConfig;
    }
}

/**
 * Extract complete workbook data with row-based batching
 */
async function getFullWorkbookData() {
    if (!workbookConfig) {
        await fetchWorkbookConfig();
    }
    
    return await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        
        workbook.load('name');
        worksheets.load('items/name');
        await context.sync();
        
        const workbookName = workbook.name || 'Untitled Workbook';
        const allBatches = [];
        
        for (let worksheet of worksheets.items) {
            const usedRange = worksheet.getUsedRange();
            if (!usedRange) continue;
            
            usedRange.load(['address', 'rowCount', 'columnCount']);
            await context.sync();
            
            if (usedRange.isNullObject) continue;
            
            if (usedRange.rowCount > workbookConfig.MAX_TOTAL_ROWS) {
                console.warn(`Worksheet ${worksheet.name} exceeds row limit`);
                continue;
            }
            
            if (usedRange.columnCount > workbookConfig.MAX_TOTAL_COLS) {
                console.warn(`Worksheet ${worksheet.name} exceeds column limit`);
                continue;
            }
            
            const worksheetBatches = await createWorksheetBatches(
                worksheet, 
                usedRange, 
                workbookConfig.MAX_ROWS_PER_BATCH
            );
            
            allBatches.push(...worksheetBatches);
        }
        
        return {
            workbookName: workbookName,
            batches: allBatches,
            totalWorksheets: worksheets.items.length
        };
    });
}

/**
 * Create batches for a single worksheet
 */
async function createWorksheetBatches(worksheet, usedRange, maxRowsPerBatch) {
    return await Excel.run(async (context) => {
        const batches = [];
        const totalRows = usedRange.rowCount;
        const address = usedRange.address;
        
        const [startCell, endCell] = address.split(':');
        const startRow = parseInt(startCell.match(/\d+/)[0]);
        const startCol = startCell.match(/[A-Z]+/)[0];
        const endCol = endCell.match(/[A-Z]+/)[0];
        
        for (let i = 0; i < totalRows; i += maxRowsPerBatch) {
            const batchStartRow = startRow + i;
            const batchEndRow = Math.min(startRow + i + maxRowsPerBatch - 1, startRow + totalRows - 1);
            const batchAddress = `${startCol}${batchStartRow}:${endCol}${batchEndRow}`;
            
            const batchRange = worksheet.getRange(batchAddress);
            batchRange.load(['values', 'formulas', 'address']);
            await context.sync();
            
            batches.push({
                worksheetName: worksheet.name,
                data: {
                    address: batchRange.address,
                    values: batchRange.values,
                    formulas: batchRange.formulas,
                    batchIndex: Math.floor(i / maxRowsPerBatch),
                    rowStart: batchStartRow,
                    rowEnd: batchEndRow
                }
            });
        }
        
        return batches;
    });
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
            workbookName: currentWorkbookName
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

