# Excel AI Chat Add-in

A minimal Office.js Excel task pane add-in with AI chat interface built with vanilla JavaScript.

## Features

- **AI Chat Interface**: Clean, modern chat UI for interacting with AI about Excel data
- **Real-time Messaging**: Send questions and receive AI responses about your spreadsheet
- **Excel Context Integration**: AI has access to your selected ranges and worksheet data
- **Minimal Design**: Focused on simplicity and usability within Excel task pane
- **Backend Ready**: Structured for easy integration with your AI backend service

## Project Structure

```
├── manifest.xml              # Office Add-in manifest
├── src/
│   └── taskpane/
│       ├── taskpane.html     # Chat interface UI
│       ├── taskpane.css      # Chat styling and responsive design
│       └── taskpane.js       # Chat logic and Office.js integration
├── assets/                   # Icons and static assets
├── package.json              # Node.js dependencies
└── webpack.config.js         # Build configuration
```

## Prerequisites

- Node.js (latest LTS version)
- Microsoft 365 subscription with Excel
- Visual Studio Code (recommended)

## Getting Started

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Start development server:**
   ```bash
   npm start
   ```
   This will start a local HTTPS server at `https://localhost:3000`

3. **Build for production:**
   ```bash
   npm run build
   ```

## Testing the Chat Interface

### Method 1: Sideload in Excel Online
1. Start the development server: `npm start`
2. Go to Excel Online and create a workbook
3. Insert > Office Add-ins > Upload My Add-in
4. Upload the `manifest.xml` file
5. Open the chat interface from the ribbon
6. Try asking questions like:
   - "What data is in my selected range?"
   - "Help me create a formula"
   - "Analyze this data"

### Method 2: Sideload in Excel Desktop
1. Follow Microsoft's instructions for sideloading add-ins
2. Use the `manifest.xml` file to register the add-in

## Backend Integration

The chat interface is ready for backend integration. Update the API_URL in `taskpane.js`:

```javascript
const API_URL = 'https://your-backend-api.com/api/chat';
```

### Expected API Format:
```javascript
// Request
POST /api/chat
{
  "message": "User's question",
  "context": {
    "selectedRange": { "address": "A1:B10", "values": [...] },
    "worksheet": { "name": "Sheet1" }
  }
}

// Response
{
  "response": "AI's answer"
}
```

## Key Technologies

- **Vanilla JavaScript**: No frameworks for maximum simplicity
- **Office.js**: Microsoft's JavaScript API for Excel integration
- **Fetch API**: For backend communication
- **CSS Grid/Flexbox**: Modern, responsive chat layout
- **Webpack**: Development server with HTTPS support

## Chat Features

- **Message Bubbles**: User (right, blue) vs AI (left, gray)
- **Auto-scroll**: Always shows latest messages
- **Loading States**: Visual feedback during AI responses
- **Error Handling**: Graceful fallbacks when backend unavailable
- **Keyboard Support**: Enter to send, auto-focus input
- **Responsive Design**: Works in narrow task pane
- **Excel Context**: Sends selected range data to AI

## 🌐 **Live App**
**https://mavencat.github.io/frontend/taskpane.html**

## Resources

- [Office Add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API reference](https://docs.microsoft.com/en-us/javascript/api/excel)
- [Office Add-in samples](https://github.com/OfficeDev/Office-Add-in-samples)