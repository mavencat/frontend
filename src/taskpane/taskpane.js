/* global console, document, Excel, Office */

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
 * Initialize chat interface
 */
function initializeChat() {
    const messageInput = document.getElementById('message-input');
    const sendButton = document.getElementById('send-button');
    
    // Event listeners
    sendButton.addEventListener('click', handleSendMessage);
    messageInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSendMessage();
        }
    });
    
    // Focus input
    messageInput.focus();
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
 * Set loading state
 * @param {boolean} isLoading - Loading state
 */
function setLoadingState(isLoading) {
    chatState.isLoading = isLoading;
    const sendButton = document.getElementById('send-button');
    const messageInput = document.getElementById('message-input');
    const status = document.getElementById('status');
    
    sendButton.disabled = isLoading;
    messageInput.disabled = isLoading;
    
    if (isLoading) {
        status.className = 'status-typing';
        status.innerHTML = '<span class="typing-indicator">AI is thinking</span>';
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

/**
 * Get AI response from backend API
 * @param {string} userMessage - User's message
 * @returns {Promise<string>} AI response
 */
async function getAIResponse(userMessage) {
    // TODO: Replace with your actual backend API URL
    const API_URL = 'http://localhost:8000/api/chat'; // Example backend URL
    
    try {
        // Get Excel context for AI
        const excelContext = await getExcelContext();
        
        // Prepare request payload
        const payload = {
            message: userMessage,
            context: excelContext,
            timestamp: new Date().toISOString()
        };
        
        // Make API call to backend
        const response = await fetch(API_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(payload)
        });
        
        if (!response.ok) {
            throw new Error(`API request failed: ${response.status}`);
        }
        
        const data = await response.json();
        return data.response || 'Sorry, I didn\'t receive a proper response.';
        
    } catch (error) {
        console.error('API Error:', error);
        
        // Fallback to placeholder responses when API is not available
        return getPlaceholderResponse(userMessage);
    }
}

/**
 * Get placeholder response (used when API is not available)
 * @param {string} userMessage - User's message
 * @returns {Promise<string>} Placeholder response
 */
async function getPlaceholderResponse(userMessage) {
    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 1000 + Math.random() * 2000));
    
    // Placeholder responses for testing
    const responses = [
        "I can help you analyze your Excel data. What specific information are you looking for?",
        "Let me help you with that. Could you select the range of cells you'd like me to examine?",
        "I understand you want to work with your spreadsheet data. What kind of analysis would be most helpful?",
        "That's a great question! I can assist with formulas, data analysis, and insights from your Excel workbook.",
        "I'm here to help with your Excel data. What would you like to know or accomplish?"
    ];
    
    // Simple response selection based on keywords
    const lowerMessage = userMessage.toLowerCase();
    if (lowerMessage.includes('formula') || lowerMessage.includes('calculate')) {
        return "I can help you create formulas! What calculation do you need to perform?";
    }
    if (lowerMessage.includes('chart') || lowerMessage.includes('graph')) {
        return "Charts are great for visualizing data! What type of chart would work best for your data?";
    }
    if (lowerMessage.includes('data') || lowerMessage.includes('analyze')) {
        return "I'd be happy to help analyze your data. Could you tell me more about what insights you're looking for?";
    }
    
    // Random response for other messages
    return responses[Math.floor(Math.random() * responses.length)];
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
            
            selectedRange.load(['address', 'values', 'rowCount', 'columnCount']);
            worksheet.load(['name']);
            
            await context.sync();
            
            return {
                selectedRange: {
                    address: selectedRange.address,
                    values: selectedRange.values,
                    rowCount: selectedRange.rowCount,
                    columnCount: selectedRange.columnCount
                },
                worksheet: {
                    name: worksheet.name
                }
            };
        });
    } catch (error) {
        console.error('Error getting Excel context:', error);
        return null;
    }
}