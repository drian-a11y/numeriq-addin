# Numeriq Quick Start Guide

## Installation & Setup

1. **Install Node.js dependencies:**
```bash
cd numeriq-addin
npm install
```

2. **Start the development server:**
```bash
npm start
```
This will start the webpack dev server on https://localhost:3000

3. **Load the add-in in Excel:**
   - Open Excel (Desktop or Online)
   - Go to **Insert** > **My Add-ins** > **Upload My Add-in**
   - Browse and select `manifest.xml` from the project root
   - The Numeriq add-in will appear in the Home tab

## Project Structure

```
numeriq-addin/
├── src/
│   ├── taskpane/
│   │   ├── components/
│   │   │   ├── App.tsx                    # Main app component
│   │   │   ├── FormulaExplorer.tsx        # Formula exploration UI
│   │   │   ├── ComparisonTool.tsx         # Spreadsheet comparison UI
│   │   │   ├── FormulaMapTool.tsx         # Formula mapping UI
│   │   │   ├── DependentsTracer.tsx       # Dependencies tracing UI
│   │   │   ├── CalculationFlowTool.tsx    # Flow analysis UI
│   │   │   └── Settings.tsx               # Settings UI
│   │   ├── taskpane.tsx                   # Entry point
│   │   ├── taskpane.html                  # HTML template
│   │   └── taskpane.css                   # Styles
│   ├── commands/
│   │   ├── commands.ts                    # Ribbon commands
│   │   └── commands.html
│   ├── utils/
│   │   ├── formulaParser.ts               # Parse formulas into trees
│   │   ├── excelHelper.ts                 # Excel API helpers
│   │   ├── comparisonHelper.ts            # Comparison logic
│   │   ├── formulaMapper.ts               # Formula mapping logic
│   │   ├── calculationFlow.ts             # Flow analysis logic
│   │   └── keyboardShortcuts.ts           # Keyboard shortcuts
│   └── types/
│       └── office.d.ts                    # Type declarations
├── assets/                                 # Icons (add your own)
├── manifest.xml                           # Add-in manifest
├── package.json
├── webpack.config.js
├── tsconfig.json
└── .babelrc
```

## Features Implemented

### 1. Formula Explorer (Ctrl+Q)
- Parse and visualize formula structure as a tree
- Navigate to precedents by clicking
- Edit formulas with F2
- Highlight active branches for IF, IFS, CHOOSE, SWITCH
- Show target locations for VLOOKUP, INDEX, OFFSET, INDIRECT

### 2. Spreadsheet Comparison (Ctrl+Shift+S/C)
- Set reference and comparator sheets
- Compare formulas or values
- Detect alignment issues
- Navigate through difference blocks
- Copy formulas between sheets
- Merge content

### 3. Formula Map (Ctrl+Shift+M)
- Color-code formulas by pattern
- Identify unique vs copied formulas
- Highlight external references
- Show hardcoded values
- Customizable color schemes
- Statistics display

### 4. Dependents Tracer (Ctrl+Shift+Q)
- Trace precedents and dependents
- Group by unique formulas
- Visual highlighting
- Navigate through results
- Multi-cell selection support

### 5. Calculation Flow (Ctrl+Shift+F)
- Analyze inputs, calculations, outputs
- Identify orphan formulas
- Inflows/outflows analysis
- Apply color-coded formatting
- Workbook/worksheet/range scope

### 6. Settings
- Customize keyboard shortcuts
- Configure explorer options
- Manage navigation history
- About information

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+Q | Explore Formula |
| Ctrl+Backspace | Navigate Back |
| Ctrl+Shift+S | Set as Reference |
| Ctrl+Shift+C | Set as Comparator |
| Ctrl+Shift+M | Toggle Formula Map |
| Ctrl+Shift+Q | Trace Dependents |
| Ctrl+Shift+F | Calculation Flow |
| F2 | Edit Formula |
| Enter | Apply Changes |
| Esc | Cancel/Close |

## Development Commands

```bash
# Start development server
npm start

# Build for production
npm run build

# Development with auto-reload
npm run dev
```

## Known Limitations

1. **INDIRECT and OFFSET**: These functions create dynamic references that cannot be traced statically
2. **External Workbooks**: References to closed workbooks may not be fully resolved
3. **Array Formulas**: Complex array formulas may require special handling
4. **Performance**: Large worksheets (>10,000 cells) may experience slower analysis

## Troubleshooting

### Add-in not loading
- Ensure the dev server is running on https://localhost:3000
- Check that Excel trusts the localhost certificate
- Clear Office cache: Delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`

### TypeScript errors
- Run `npm install` to ensure all dependencies are installed
- Check that @types/office-js is installed

### Build errors
- Clear node_modules and reinstall: `rm -rf node_modules && npm install`
- Clear webpack cache: `rm -rf dist`

## Next Steps

1. **Add Icons**: Replace placeholder PNG files in `/assets` with actual icons
2. **Testing**: Test with various Excel formulas and workbooks
3. **Optimization**: Profile and optimize for large worksheets
4. **Documentation**: Add inline code documentation
5. **Error Handling**: Enhance error messages and user feedback

## Support

For issues or questions, refer to:
- Office Add-ins Documentation: https://docs.microsoft.com/office/dev/add-ins
- Office.js API Reference: https://docs.microsoft.com/javascript/api/excel
