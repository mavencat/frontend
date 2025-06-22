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
        
        // Check backend health on load
        checkBackendHealth();
        
        // Set up periodic health checks (every 5 minutes)
        setInterval(checkBackendHealth, 5 * 60 * 1000);
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

// Production backend configuration
const API_CONFIG = {
    BASE_URL: 'https://backend-ulfc334oda-ew.a.run.app',
    ENDPOINTS: {
        CHAT: '/chat',
        HEALTH: '/health'
    },
    TIMEOUT: 30000 // 30 seconds timeout
};

// Build full URLs
const API_URLS = {
    chat: `${API_CONFIG.BASE_URL}${API_CONFIG.ENDPOINTS.CHAT}`,
    health: `${API_CONFIG.BASE_URL}${API_CONFIG.ENDPOINTS.HEALTH}`
};

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
            timestamp: new Date().toISOString()
        };
        
        console.log('Sending request to:', API_URLS.chat);
        console.log('Payload:', payload);
        
        // Make API call with timeout
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), API_CONFIG.TIMEOUT);
        
        const response = await fetch(API_URLS.chat, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify(payload),
            signal: controller.signal
        });
        
        clearTimeout(timeoutId);
        
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
        
        let errorMessage = 'Sorry, I encountered an error. ';
        
        if (error.name === 'AbortError') {
            errorMessage += 'The request timed out. Please try again.';
        } else if (error.message.includes('Failed to fetch')) {
            errorMessage += 'Unable to connect to the server. Please check your internet connection.';
        } else {
            errorMessage += error.message;
        }
        
        throw new Error(errorMessage);
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

/**
 * Check backend health status
 * @returns {Promise<boolean>} Health status
 */
async function checkBackendHealth() {
    try {
        console.log('Checking backend health...');
        
        const response = await fetch(API_URLS.health, {
            method: 'GET',
            headers: {
                'Accept': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`Health check failed: ${response.status}`);
        }
        
        const health = await response.json();
        console.log('Backend health check passed:', health);
        
        // Update UI to show backend status
        const statusElement = document.getElementById('backendStatus');
        if (statusElement) {
            statusElement.textContent = '✅ Backend connected';
            statusElement.className = 'status-healthy';
        }
        
        return true;
        
    } catch (error) {
        console.error('Backend health check failed:', error);
        
        // Update UI to show backend status
        const statusElement = document.getElementById('backendStatus');
        if (statusElement) {
            statusElement.textContent = '❌ Backend offline';
            statusElement.className = 'status-error';
        }
        
        return false;
    }
}