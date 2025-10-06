import * as React from 'react';
import { ExcelHelper, CellInfo } from '../../utils/excelHelper';

/* global Excel */

export interface DependentsTracerState {
  selectedRange: string;
  traceMode: 'precedents' | 'dependents';
  results: CellInfo[];
  groupedResults: Map<string, CellInfo[]>;
  selectedGroup: string | null;
  highlightColors: {
    selected: string;
    precedents: string;
    dependents: string;
  };
}

export class DependentsTracer extends React.Component<{}, DependentsTracerState> {
  constructor(props: {}) {
    super(props);

    this.state = {
      selectedRange: '',
      traceMode: 'dependents',
      results: [],
      groupedResults: new Map(),
      selectedGroup: null,
      highlightColors: {
        selected: '#FFB6C1',  // Pink
        precedents: '#ADD8E6', // Light blue
        dependents: '#90EE90'  // Light green
      }
    };
  }

  componentDidMount() {
    this.loadSelectedRange();
  }

  loadSelectedRange = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('address');
        await context.sync();

        this.setState({ selectedRange: range.address });
      });
    } catch (error) {
      console.error('Error loading selected range:', error);
    }
  };

  traceDependents = async () => {
    const { selectedRange } = this.state;

    if (!selectedRange) {
      alert('Please select a range first');
      return;
    }

    try {
      await Excel.run(async (context) => {
        const [sheetPart, addressPart] = selectedRange.includes('!')
          ? selectedRange.split('!')
          : [null, selectedRange];

        const sheet = sheetPart
          ? context.workbook.worksheets.getItem(sheetPart)
          : context.workbook.worksheets.getActiveWorksheet();

        sheet.load('name');
        await context.sync();

        const dependents = await ExcelHelper.getDirectDependents(
          context,
          addressPart,
          sheet.name
        );

        // Group by unique formulas
        const grouped = this.groupByFormula(dependents);

        // Highlight the selected range
        await ExcelHelper.highlightRange(
          context,
          addressPart,
          this.state.highlightColors.selected,
          sheet.name
        );

        // Highlight dependents
        for (const dependent of dependents) {
          await ExcelHelper.highlightRange(
            context,
            dependent.address,
            this.state.highlightColors.dependents,
            dependent.sheet
          );
        }

        this.setState({
          results: dependents,
          groupedResults: grouped,
          traceMode: 'dependents'
        });
      });
    } catch (error) {
      console.error('Error tracing dependents:', error);
    }
  };

  tracePrecedents = async () => {
    const { selectedRange } = this.state;

    if (!selectedRange) {
      alert('Please select a range first');
      return;
    }

    try {
      await Excel.run(async (context) => {
        const [sheetPart, addressPart] = selectedRange.includes('!')
          ? selectedRange.split('!')
          : [null, selectedRange];

        const sheet = sheetPart
          ? context.workbook.worksheets.getItem(sheetPart)
          : context.workbook.worksheets.getActiveWorksheet();

        sheet.load('name');
        await context.sync();

        const precedents = await ExcelHelper.getDirectPrecedents(
          context,
          addressPart,
          sheet.name
        );

        // Convert precedents to CellInfo format
        const precedentCells: CellInfo[] = precedents.map(p => ({
          address: p.address,
          formula: '',
          value: p.value,
          sheet: p.sheet,
          workbook: p.workbook
        }));

        // Highlight the selected range
        await ExcelHelper.highlightRange(
          context,
          addressPart,
          this.state.highlightColors.selected,
          sheet.name
        );

        // Highlight precedents
        for (const precedent of precedents) {
          await ExcelHelper.highlightRange(
            context,
            precedent.address,
            this.state.highlightColors.precedents,
            precedent.sheet
          );
        }

        const grouped = this.groupByFormula(precedentCells);

        this.setState({
          results: precedentCells,
          groupedResults: grouped,
          traceMode: 'precedents'
        });
      });
    } catch (error) {
      console.error('Error tracing precedents:', error);
    }
  };

  groupByFormula = (cells: CellInfo[]): Map<string, CellInfo[]> => {
    const grouped = new Map<string, CellInfo[]>();

    for (const cell of cells) {
      const key = cell.formula || 'No Formula';
      
      if (!grouped.has(key)) {
        grouped.set(key, []);
      }
      
      grouped.get(key)!.push(cell);
    }

    return grouped;
  };

  navigateToGroup = async (groupKey: string) => {
    const { groupedResults } = this.state;
    const cells = groupedResults.get(groupKey);

    if (!cells || cells.length === 0) return;

    this.setState({ selectedGroup: groupKey });

    try {
      await Excel.run(async (context) => {
        const firstCell = cells[0];
        await ExcelHelper.navigateToCell(context, firstCell.address, firstCell.sheet);
      });
    } catch (error) {
      console.error('Error navigating to group:', error);
    }
  };

  navigateToCell = async (cell: CellInfo) => {
    try {
      await Excel.run(async (context) => {
        await ExcelHelper.navigateToCell(context, cell.address, cell.sheet);
      });
    } catch (error) {
      console.error('Error navigating to cell:', error);
    }
  };

  clearHighlights = async () => {
    try {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load('items');
        await context.sync();

        for (const sheet of sheets.items) {
          const usedRange = sheet.getUsedRange();
          usedRange.format.fill.clear();
        }

        await context.sync();
      });
    } catch (error) {
      console.error('Error clearing highlights:', error);
    }
  };

  render() {
    const { selectedRange, traceMode, results, groupedResults, selectedGroup, highlightColors } = this.state;

    return (
      <div className="dependents-tracer">
        <div className="tool-header">
          <h2>Trace Dependencies</h2>
          <p className="tool-description">
            Trace precedents and dependents of multiple cells simultaneously
          </p>
        </div>

        <div className="range-selection">
          <h3>Selected Range</h3>
          <div className="range-display">
            <input
              type="text"
              value={selectedRange}
              readOnly
              className="range-input"
              placeholder="No range selected"
            />
            <button className="btn btn-secondary" onClick={this.loadSelectedRange}>
              Refresh
            </button>
          </div>
        </div>

        <div className="trace-actions">
          <button className="btn btn-primary" onClick={this.traceDependents}>
            Trace Dependents (Ctrl+Shift+Q)
          </button>
          <button className="btn btn-primary" onClick={this.tracePrecedents}>
            Trace Precedents (Ctrl+Q)
          </button>
          <button className="btn btn-secondary" onClick={this.clearHighlights}>
            Clear Highlights
          </button>
        </div>

        <div className="highlight-legend">
          <div className="legend-item">
            <span className="legend-color" style={{ backgroundColor: highlightColors.selected }}></span>
            <span className="legend-label">Selected Range</span>
          </div>
          <div className="legend-item">
            <span className="legend-color" style={{ backgroundColor: highlightColors.precedents }}></span>
            <span className="legend-label">Precedents</span>
          </div>
          <div className="legend-item">
            <span className="legend-color" style={{ backgroundColor: highlightColors.dependents }}></span>
            <span className="legend-label">Dependents</span>
          </div>
        </div>

        {results.length > 0 && (
          <div className="trace-results">
            <div className="results-header">
              <h3>
                {traceMode === 'dependents' ? 'Dependents' : 'Precedents'} Found: {results.length}
              </h3>
              <p className="results-summary">
                Grouped into {groupedResults.size} unique {groupedResults.size === 1 ? 'block' : 'blocks'}
              </p>
            </div>

            <div className="grouped-results">
              {Array.from(groupedResults.entries()).map(([groupKey, cells], index) => (
                <div
                  key={index}
                  className={`result-group ${selectedGroup === groupKey ? 'selected' : ''}`}
                  onClick={() => this.navigateToGroup(groupKey)}
                >
                  <div className="group-header">
                    <h4>Block {index + 1}</h4>
                    <span className="cell-count">{cells.length} {cells.length === 1 ? 'cell' : 'cells'}</span>
                  </div>
                  
                  {groupKey !== 'No Formula' && (
                    <div className="group-formula">
                      <code>{groupKey}</code>
                    </div>
                  )}

                  <div className="group-cells">
                    {cells.map((cell, cellIndex) => (
                      <div
                        key={cellIndex}
                        className="cell-item"
                        onClick={(e) => {
                          e.stopPropagation();
                          this.navigateToCell(cell);
                        }}
                      >
                        <span className="cell-address">
                          {cell.sheet}!{cell.address}
                        </span>
                        <span className="cell-value">= {String(cell.value)}</span>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {results.length === 0 && (
          <div className="empty-state">
            <p>No {traceMode} found for the selected range</p>
          </div>
        )}

        <div className="tool-help">
          <h4>How to Use</h4>
          <ul>
            <li>Select a range of cells in Excel</li>
            <li>Click "Trace Dependents" to find cells that reference the selection</li>
            <li>Click "Trace Precedents" to find cells referenced by the selection</li>
            <li>Navigate through grouped results by clicking on blocks</li>
            <li>Use <kbd>Home</kbd> to return to the original selection</li>
            <li>Press <kbd>Ctrl+Shift+Q</kbd> or <kbd>Ctrl+Q</kbd> on selected elements to trace further</li>
          </ul>
        </div>
      </div>
    );
  }
}
