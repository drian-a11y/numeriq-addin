import * as React from 'react';
import { ComparisonHelper, ComparisonResult, ComparisonOptions, DifferenceBlock } from '../../utils/comparisonHelper';

/* global Excel */

export interface ComparisonToolState {
  referenceSheet: string;
  comparatorSheets: string[];
  availableSheets: string[];
  comparisonMode: 'workbook' | 'worksheet' | 'range';
  options: ComparisonOptions;
  results: ComparisonResult | null;
  selectedBlock: DifferenceBlock | null;
  currentBlockIndex: number;
  isComparing: boolean;
}

export class ComparisonTool extends React.Component<{}, ComparisonToolState> {
  constructor(props: {}) {
    super(props);

    this.state = {
      referenceSheet: '',
      comparatorSheets: [],
      availableSheets: [],
      comparisonMode: 'worksheet',
      options: {
        compareFormulas: true,
        ignoreInputs: false,
        detectAlignment: true
      },
      results: null,
      selectedBlock: null,
      currentBlockIndex: 0,
      isComparing: false
    };
  }

  componentDidMount() {
    this.loadAvailableSheets();
  }

  loadAvailableSheets = async () => {
    try {
      await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load('items/name');
        await context.sync();

        const sheetNames = worksheets.items.map(ws => ws.name);
        this.setState({ availableSheets: sheetNames });
      });
    } catch (error) {
      console.error('Error loading sheets:', error);
    }
  };

  setAsReference = async () => {
    try {
      await Excel.run(async (context) => {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load('name');
        await context.sync();

        this.setState({ referenceSheet: activeSheet.name });
      });
    } catch (error) {
      console.error('Error setting reference:', error);
    }
  };

  setAsComparator = async () => {
    try {
      await Excel.run(async (context) => {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load('name');
        await context.sync();

        this.setState(prevState => ({
          comparatorSheets: [...prevState.comparatorSheets, activeSheet.name]
        }));
      });
    } catch (error) {
      console.error('Error setting comparator:', error);
    }
  };

  removeComparator = (sheetName: string) => {
    this.setState(prevState => ({
      comparatorSheets: prevState.comparatorSheets.filter(s => s !== sheetName)
    }));
  };

  compareSheets = async () => {
    const { referenceSheet, comparatorSheets, options } = this.state;

    if (!referenceSheet || comparatorSheets.length === 0) {
      alert('Please set both reference and comparator sheets');
      return;
    }

    this.setState({ isComparing: true });

    try {
      await Excel.run(async (context) => {
        // For now, compare with the first comparator
        const result = await ComparisonHelper.compareWorksheets(
          context,
          referenceSheet,
          comparatorSheets[0],
          options
        );

        this.setState({
          results: result,
          currentBlockIndex: 0,
          selectedBlock: result.differences[0] || null,
          isComparing: false
        });
      });
    } catch (error) {
      console.error('Error comparing sheets:', error);
      this.setState({ isComparing: false });
    }
  };

  navigateToBlock = async (block: DifferenceBlock, index: number) => {
    this.setState({ selectedBlock: block, currentBlockIndex: index });

    try {
      await Excel.run(async (context) => {
        const { referenceSheet } = this.state;
        const sheet = context.workbook.worksheets.getItem(referenceSheet);
        
        const address = `${this.getColumnLetter(block.startCol + 1)}${block.startRow + 1}`;
        const range = sheet.getRange(address);
        range.select();
        
        await context.sync();
      });
    } catch (error) {
      console.error('Error navigating to block:', error);
    }
  };

  nextBlock = () => {
    const { results, currentBlockIndex } = this.state;
    if (results && currentBlockIndex < results.differences.length - 1) {
      const nextIndex = currentBlockIndex + 1;
      this.navigateToBlock(results.differences[nextIndex], nextIndex);
    }
  };

  previousBlock = () => {
    const { currentBlockIndex, results } = this.state;
    if (currentBlockIndex > 0) {
      const prevIndex = currentBlockIndex - 1;
      this.navigateToBlock(results!.differences[prevIndex], prevIndex);
    }
  };

  copyToRight = async () => {
    const { selectedBlock, referenceSheet, comparatorSheets } = this.state;
    if (!selectedBlock || comparatorSheets.length === 0) return;

    try {
      await Excel.run(async (context) => {
        for (const cell of selectedBlock.referenceCells) {
          await ComparisonHelper.copyFormula(
            context,
            referenceSheet,
            comparatorSheets[0],
            cell.address
          );
        }
      });

      // Re-compare after copying
      await this.compareSheets();
    } catch (error) {
      console.error('Error copying to right:', error);
    }
  };

  copyToLeft = async () => {
    const { selectedBlock, referenceSheet, comparatorSheets } = this.state;
    if (!selectedBlock || comparatorSheets.length === 0) return;

    try {
      await Excel.run(async (context) => {
        for (const cell of selectedBlock.comparatorCells) {
          await ComparisonHelper.copyFormula(
            context,
            comparatorSheets[0],
            referenceSheet,
            cell.address
          );
        }
      });

      // Re-compare after copying
      await this.compareSheets();
    } catch (error) {
      console.error('Error copying to left:', error);
    }
  };

  alignWorksheets = async () => {
    const { results, referenceSheet } = this.state;
    if (!results || !results.alignmentNeeded) return;

    try {
      await Excel.run(async (context) => {
        await ComparisonHelper.insertAlignmentRows(
          context,
          referenceSheet,
          results.insertedRows
        );
      });

      // Re-compare after alignment
      await this.compareSheets();
    } catch (error) {
      console.error('Error aligning worksheets:', error);
    }
  };

  removeAlignmentRows = async () => {
    const { referenceSheet } = this.state;

    try {
      await Excel.run(async (context) => {
        await ComparisonHelper.removeAlignmentRows(context, referenceSheet);
      });
    } catch (error) {
      console.error('Error removing alignment rows:', error);
    }
  };

  getColumnLetter = (col: number): string => {
    let column = '';
    while (col > 0) {
      const remainder = (col - 1) % 26;
      column = String.fromCharCode(65 + remainder) + column;
      col = Math.floor((col - 1) / 26);
    }
    return column;
  };

  render() {
    const {
      referenceSheet,
      comparatorSheets,
      availableSheets,
      options,
      results,
      selectedBlock,
      currentBlockIndex,
      isComparing
    } = this.state;

    return (
      <div className="comparison-tool">
        <div className="tool-header">
          <h2>Spreadsheet Comparison</h2>
        </div>

        <div className="comparison-setup">
          <div className="setup-section">
            <h3>Reference</h3>
            <div className="reference-display">
              {referenceSheet ? (
                <div className="selected-sheet">
                  <span className="sheet-name">{referenceSheet}</span>
                  <button className="btn-icon" onClick={() => this.setState({ referenceSheet: '' })}>
                    ✕
                  </button>
                </div>
              ) : (
                <p className="placeholder">No reference set</p>
              )}
            </div>
            <button className="btn btn-primary" onClick={this.setAsReference}>
              Set as Reference (Ctrl+Shift+S)
            </button>
          </div>

          <div className="setup-section">
            <h3>Comparators</h3>
            <div className="comparators-list">
              {comparatorSheets.length > 0 ? (
                comparatorSheets.map((sheet, index) => (
                  <div key={index} className="selected-sheet">
                    <span className="sheet-name">{sheet}</span>
                    <button className="btn-icon" onClick={() => this.removeComparator(sheet)}>
                      ✕
                    </button>
                  </div>
                ))
              ) : (
                <p className="placeholder">No comparators set</p>
              )}
            </div>
            <button className="btn btn-primary" onClick={this.setAsComparator}>
              Add Comparator (Ctrl+Shift+C)
            </button>
          </div>
        </div>

        <div className="comparison-options">
          <h3>Options</h3>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={options.compareFormulas}
              onChange={(e) => this.setState({
                options: { ...options, compareFormulas: e.target.checked }
              })}
            />
            Compare Formulas (vs Values)
          </label>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={options.ignoreInputs}
              onChange={(e) => this.setState({
                options: { ...options, ignoreInputs: e.target.checked }
              })}
              disabled={!options.compareFormulas}
            />
            Ignore Differences in Inputs
          </label>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={options.detectAlignment}
              onChange={(e) => this.setState({
                options: { ...options, detectAlignment: e.target.checked }
              })}
            />
            Detect Alignment Issues
          </label>
        </div>

        <div className="comparison-actions">
          <button
            className="btn btn-primary btn-large"
            onClick={this.compareSheets}
            disabled={!referenceSheet || comparatorSheets.length === 0 || isComparing}
          >
            {isComparing ? 'Comparing...' : 'Compare'}
          </button>
        </div>

        {results && (
          <div className="comparison-results">
            <div className="results-header">
              <h3>Results</h3>
              <div className="results-summary">
                <span className="difference-count">
                  {results.totalDifferences} differences found in {results.differences.length} blocks
                </span>
              </div>
            </div>

            {results.alignmentNeeded && (
              <div className="alignment-warning">
                <p>⚠️ Alignment issues detected</p>
                <button className="btn btn-secondary" onClick={this.alignWorksheets}>
                  Align Worksheets
                </button>
                <button className="btn btn-secondary" onClick={this.removeAlignmentRows}>
                  Remove Alignment Rows
                </button>
              </div>
            )}

            {selectedBlock && (
              <div className="block-viewer">
                <div className="block-navigation">
                  <button
                    className="btn btn-icon"
                    onClick={this.previousBlock}
                    disabled={currentBlockIndex === 0}
                  >
                    ←
                  </button>
                  <span className="block-counter">
                    Block {currentBlockIndex + 1} of {results.differences.length}
                  </span>
                  <button
                    className="btn btn-icon"
                    onClick={this.nextBlock}
                    disabled={currentBlockIndex === results.differences.length - 1}
                  >
                    →
                  </button>
                </div>

                <div className="block-comparison">
                  <div className="comparison-side">
                    <h4>Reference: {referenceSheet}</h4>
                    <div className="cell-list">
                      {selectedBlock.referenceCells.map((cell, index) => (
                        <div key={index} className={`cell-item ${cell.isDifferent ? 'different' : ''}`}>
                          <span className="cell-address">{cell.address}</span>
                          <code className="cell-formula">{cell.formula || String(cell.value)}</code>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="comparison-actions-middle">
                    <button className="btn btn-secondary" onClick={this.copyToRight} title="Copy to Right (Ctrl+R)">
                      →
                    </button>
                    <button className="btn btn-secondary" onClick={this.copyToLeft} title="Copy to Left (Ctrl+L)">
                      ←
                    </button>
                  </div>

                  <div className="comparison-side">
                    <h4>Comparator: {comparatorSheets[0]}</h4>
                    <div className="cell-list">
                      {selectedBlock.comparatorCells.map((cell, index) => (
                        <div key={index} className={`cell-item ${cell.isDifferent ? 'different' : ''}`}>
                          <span className="cell-address">{cell.address}</span>
                          <code className="cell-formula">{cell.formula || String(cell.value)}</code>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        <div className="tool-help">
          <h4>Keyboard Shortcuts</h4>
          <ul>
            <li><kbd>Ctrl+Shift+S</kbd> - Set as Reference</li>
            <li><kbd>Ctrl+Shift+C</kbd> - Set as Comparator</li>
            <li><kbd>Ctrl+R</kbd> - Copy to Right</li>
            <li><kbd>Ctrl+L</kbd> - Copy to Left</li>
            <li><kbd>Ctrl+Tab</kbd> - Switch sides</li>
            <li><kbd>Alt+R</kbd> - Re-compare</li>
            <li><kbd>Alt+X</kbd> - Export report</li>
          </ul>
        </div>
      </div>
    );
  }
}
