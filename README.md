# Excel Task Pane Add-in

A minimal Office.js Excel task pane add-in built with JavaScript.

## Features

- Highlight selected Excel range in yellow
- Insert sample text into selected cells
- Get information about selected range (address, dimensions, values)
- Clean, responsive UI using Office UI Fabric-like styling

## Project Structure

```
├── manifest.xml              # Office Add-in manifest
├── src/
│   └── taskpane/
│       ├── taskpane.html     # Task pane UI
│       ├── taskpane.css      # Styling
│       └── taskpane.js       # Office.js API interactions
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

## Testing the Add-in

### Method 1: Sideload in Excel Online
1. Go to Excel Online
2. Open or create a workbook
3. Go to Insert > Office Add-ins > Upload My Add-in
4. Upload the `manifest.xml` file
5. The add-in will appear in the ribbon

### Method 2: Sideload in Excel Desktop
1. Follow Microsoft's instructions for sideloading add-ins in Excel desktop
2. Use the `manifest.xml` file to register the add-in

## Key Technologies

- **Office.js**: Microsoft's JavaScript API for Office applications
- **Excel JavaScript API**: Specific APIs for Excel workbook manipulation
- **Webpack**: Module bundler and development server
- **HTTPS**: Required for Office add-ins

## Next Steps

This is a minimal strawman implementation. Consider adding:

- More Excel operations (charts, tables, formatting)
- Authentication and external data sources
- Custom functions
- Ribbon buttons and commands
- Unit tests
- Production deployment setup

## Resources

- [Office Add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API reference](https://docs.microsoft.com/en-us/javascript/api/excel)
- [Office Add-in samples](https://github.com/OfficeDev/Office-Add-in-samples)