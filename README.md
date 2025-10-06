# Numeriq - Excel Add-in

Advanced Excel formula exploration and analysis tool.

## Features

### 1. Formula Explorer
- View the logical structure of any formula in a pop-up window
- Navigate easily to precedents and back
- See formula components with their values and locations
- Edit formulas directly with F2
- Support for logical functions (IF, IFS, CHOOSE, SWITCH)
- Support for reference functions (VLOOKUP, OFFSET, INDEX, INDIRECT)

### 2. Spreadsheet Comparison
- Compare workbooks, worksheets, and blocks of cells
- Detect inserted rows/columns and align automatically
- Compare formulas or values
- Navigate through differences
- Merge content between sheets
- Export comparison reports

### 3. Formula Mapping
- Apply color schemes to reveal formula patterns
- Identify unique formulas and inconsistencies
- Detect external references
- Highlight hardcoded data
- Customize color schemes

### 4. Multi-Cell Dependents Tracing
- Trace direct precedents or dependents of multiple cells
- Navigate through grouped results
- Explore calculation trees
- Visual highlighting of dependencies

### 5. Calculation Flow Analysis
- Review data flow in and out of calculation areas
- Identify inputs, calculations, and outputs
- Analyze inflows and outflows
- Apply logic-based formatting

## Installation

1. Install dependencies:
```bash
npm install
```

2. Start the development server:
```bash
npm start
```

3. Sideload the add-in in Excel:
   - Open Excel
   - Go to Insert > My Add-ins > Upload My Add-in
   - Select the `manifest.xml` file

## Development

### Build for production:
```bash
npm run build
```

### Project Structure:
```
numeriq-addin/
├── src/
│   ├── taskpane/
│   │   ├── components/     # React components
│   │   ├── taskpane.tsx    # Main entry point
│   │   ├── taskpane.html   # HTML template
│   │   └── taskpane.css    # Styles
│   ├── commands/           # Ribbon commands
│   └── utils/              # Utility modules
│       ├── formulaParser.ts
│       ├── excelHelper.ts
│       ├── comparisonHelper.ts
│       ├── formulaMapper.ts
│       ├── calculationFlow.ts
│       └── keyboardShortcuts.ts
├── assets/                 # Icons and images
├── manifest.xml           # Add-in manifest
├── package.json
├── webpack.config.js
└── tsconfig.json
```

## Keyboard Shortcuts

- **Ctrl+Q** - Explore Formula
- **Ctrl+Backspace** - Navigate Back
- **Ctrl+Shift+S** - Set as Reference
- **Ctrl+Shift+C** - Set as Comparator
- **Ctrl+Shift+M** - Toggle Formula Map
- **Ctrl+Shift+Q** - Trace Dependents
- **Ctrl+Shift+F** - Calculation Flow
- **F2** - Edit Formula
- **Enter** - Apply Changes
- **Esc** - Cancel/Close

## Technologies Used

- TypeScript
- React
- Office.js API
- Webpack
- Babel

## License

MIT
