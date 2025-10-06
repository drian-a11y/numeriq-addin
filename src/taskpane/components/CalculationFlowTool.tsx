import * as React from 'react';
import { CalculationFlowAnalyzer, FlowAnalysisResult, FlowColors, AnalysisScope } from '../../utils/calculationFlow';

/* global Excel */

export interface CalculationFlowToolState {
  scopeType: 'workbook' | 'worksheet' | 'range';
  selectedSheet: string;
  scopeRange: string;
  focusAreaEnabled: boolean;
  focusSheet: string;
  focusRange: string;
  availableSheets: string[];
  colors: FlowColors;
  results: FlowAnalysisResult | null;
  isAnalyzing: boolean;
  colorsApplied: boolean;
  activeView: 'inputs-outputs' | 'inflows-outflows';
}

export class CalculationFlowTool extends React.Component<{}, CalculationFlowToolState> {
  constructor(props: {}) {
    super(props);

    this.state = {
      scopeType: 'worksheet',
      selectedSheet: '',
      scopeRange: '',
      focusAreaEnabled: false,
      focusSheet: '',
      focusRange: '',
      availableSheets: [],
      colors: CalculationFlowAnalyzer.getDefaultColors(),
      results: null,
      isAnalyzing: false,
      colorsApplied: false,
      activeView: 'inputs-outputs'
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
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load('name');
        await context.sync();

        this.setState({
          availableSheets: sheetNames,
          selectedSheet: activeSheet.name,
          focusSheet: activeSheet.name
        });
      });
    } catch (error) {
      console.error('Error loading sheets:', error);
    }
  };

  analyzeFlow = async () => {
    const { scopeType, selectedSheet, scopeRange, focusAreaEnabled, focusSheet, focusRange, colors } = this.state;

    this.setState({ isAnalyzing: true });

    try {
      await Excel.run(async (context) => {
        const scope: AnalysisScope = {
          type: scopeType,
          sheetName: scopeType !== 'workbook' ? selectedSheet : undefined,
          rangeAddress: scopeType === 'range' ? scopeRange : undefined,
          focusArea: focusAreaEnabled ? {
            sheetName: focusSheet,
            rangeAddress: focusRange || undefined
          } : undefined
        };

        const results = await CalculationFlowAnalyzer.analyzeFlow(context, scope, colors);

        this.setState({
          results,
          isAnalyzing: false,
          activeView: focusAreaEnabled ? 'inflows-outflows' : 'inputs-outputs'
        });
      });
    } catch (error) {
      console.error('Error analyzing flow:', error);
      this.setState({ isAnalyzing: false });
    }
  };

  applyColors = async () => {
    const { results } = this.state;

    if (!results) return;

    try {
      await Excel.run(async (context) => {
        await CalculationFlowAnalyzer.applyFlowColors(context, results);
        this.setState({ colorsApplied: true });
      });
    } catch (error) {
      console.error('Error applying colors:', error);
    }
  };

  removeColors = async () => {
    const { scopeType, selectedSheet, scopeRange } = this.state;

    try {
      await Excel.run(async (context) => {
        const scope: AnalysisScope = {
          type: scopeType,
          sheetName: scopeType !== 'workbook' ? selectedSheet : undefined,
          rangeAddress: scopeType === 'range' ? scopeRange : undefined
        };

        await CalculationFlowAnalyzer.removeFlowColors(context, scope);
        this.setState({ colorsApplied: false });
      });
    } catch (error) {
      console.error('Error removing colors:', error);
    }
  };

  keepColors = () => {
    // Colors are already applied to the workbook, just update state
    this.setState({ colorsApplied: false, results: null });
  };

  render() {
    const {
      scopeType,
      selectedSheet,
      scopeRange,
      focusAreaEnabled,
      focusSheet,
      focusRange,
      availableSheets,
      colors,
      results,
      isAnalyzing,
      colorsApplied,
      activeView
    } = this.state;

    return (
      <div className="calculation-flow-tool">
        <div className="tool-header">
          <h2>Calculation Flow Analysis</h2>
          <p className="tool-description">
            Review how data flows in and out of calculation areas
          </p>
        </div>

        <div className="scope-configuration">
          <h3>Analysis Scope</h3>
          <div className="scope-type-selector">
            <label className="radio-label">
              <input
                type="radio"
                value="workbook"
                checked={scopeType === 'workbook'}
                onChange={(e) => this.setState({ scopeType: e.target.value as any })}
              />
              Entire Workbook
            </label>
            <label className="radio-label">
              <input
                type="radio"
                value="worksheet"
                checked={scopeType === 'worksheet'}
                onChange={(e) => this.setState({ scopeType: e.target.value as any })}
              />
              Worksheet
            </label>
            <label className="radio-label">
              <input
                type="radio"
                value="range"
                checked={scopeType === 'range'}
                onChange={(e) => this.setState({ scopeType: e.target.value as any })}
              />
              Range
            </label>
          </div>

          {scopeType !== 'workbook' && (
            <div className="scope-details">
              <label>
                Worksheet:
                <select
                  value={selectedSheet}
                  onChange={(e) => this.setState({ selectedSheet: e.target.value })}
                  className="select-input"
                >
                  {availableSheets.map(sheet => (
                    <option key={sheet} value={sheet}>{sheet}</option>
                  ))}
                </select>
              </label>

              {scopeType === 'range' && (
                <label>
                  Range:
                  <input
                    type="text"
                    value={scopeRange}
                    onChange={(e) => this.setState({ scopeRange: e.target.value })}
                    placeholder="e.g., A1:D10"
                    className="text-input"
                  />
                </label>
              )}
            </div>
          )}
        </div>

        <div className="focus-area-configuration">
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={focusAreaEnabled}
              onChange={(e) => this.setState({ focusAreaEnabled: e.target.checked })}
            />
            Enable Focus Area (analyze inflows/outflows)
          </label>

          {focusAreaEnabled && (
            <div className="focus-details">
              <label>
                Focus Worksheet:
                <select
                  value={focusSheet}
                  onChange={(e) => this.setState({ focusSheet: e.target.value })}
                  className="select-input"
                >
                  {availableSheets.map(sheet => (
                    <option key={sheet} value={sheet}>{sheet}</option>
                  ))}
                </select>
              </label>

              <label>
                Focus Range (optional):
                <input
                  type="text"
                  value={focusRange}
                  onChange={(e) => this.setState({ focusRange: e.target.value })}
                  placeholder="e.g., A1:D10"
                  className="text-input"
                />
              </label>
            </div>
          )}
        </div>

        <div className="flow-actions">
          <button
            className="btn btn-primary btn-large"
            onClick={this.analyzeFlow}
            disabled={isAnalyzing}
          >
            {isAnalyzing ? 'Analyzing...' : 'Analyze Flow (Ctrl+Shift+F)'}
          </button>
        </div>

        {results && (
          <div className="flow-results">
            <div className="results-header">
              <h3>Analysis Results</h3>
              <div className="view-toggle">
                <button
                  className={`btn btn-tab ${activeView === 'inputs-outputs' ? 'active' : ''}`}
                  onClick={() => this.setState({ activeView: 'inputs-outputs' })}
                >
                  Inputs - Outputs
                </button>
                {focusAreaEnabled && (
                  <button
                    className={`btn btn-tab ${activeView === 'inflows-outflows' ? 'active' : ''}`}
                    onClick={() => this.setState({ activeView: 'inflows-outflows' })}
                  >
                    Inflows - Outflows
                  </button>
                )}
              </div>
            </div>

            {activeView === 'inputs-outputs' && (
              <div className="flow-groups">
                <div className="flow-group">
                  <div className="group-header">
                    <span className="group-color" style={{ backgroundColor: colors.inputs }}></span>
                    <h4>Inputs</h4>
                    <span className="group-count">{results.inputs[0]?.cells.length || 0} cells</span>
                  </div>
                  <p className="group-description">Cells with dependents but no in-scope precedents</p>
                </div>

                <div className="flow-group">
                  <div className="group-header">
                    <span className="group-color" style={{ backgroundColor: colors.calculations }}></span>
                    <h4>Calculations</h4>
                    <span className="group-count">{results.calculations[0]?.cells.length || 0} cells</span>
                  </div>
                  <p className="group-description">Cells with both in-scope precedents and dependents</p>
                </div>

                <div className="flow-group">
                  <div className="group-header">
                    <span className="group-color" style={{ backgroundColor: colors.outputs }}></span>
                    <h4>Outputs</h4>
                    <span className="group-count">{results.outputs[0]?.cells.length || 0} cells</span>
                  </div>
                  <p className="group-description">Cells with precedents but no in-scope dependents</p>
                </div>

                <div className="flow-group">
                  <div className="group-header">
                    <span className="group-color" style={{ backgroundColor: colors.orphans }}></span>
                    <h4>Orphan Formulas</h4>
                    <span className="group-count">{results.orphans[0]?.cells.length || 0} cells</span>
                  </div>
                  <p className="group-description">Formulas without in-scope precedents or dependents</p>
                </div>
              </div>
            )}

            {activeView === 'inflows-outflows' && results.inflows && (
              <div className="flow-groups">
                <div className="flow-group">
                  <div className="group-header">
                    <span className="group-color" style={{ backgroundColor: colors.outsidePrecedents }}></span>
                    <h4>Outside - Precedents</h4>
                    <span className="group-count">{results.outsidePrecedents?.[0]?.cells.length || 0} cells</span>
                  </div>
                  <p className="group-description">Cells outside focus area that are precedents to cells inside</p>
                </div>

                <div className="flow-group">
                  <div className="group-header">
                    <span className="group-color" style={{ backgroundColor: colors.inflows }}></span>
                    <h4>Inside - Inflows</h4>
                    <span className="group-count">{results.inflows[0]?.cells.length || 0} cells</span>
                  </div>
                  <p className="group-description">Cells inside focus area that use values from outside</p>
                </div>

                <div className="flow-group">
                  <div className="group-header">
                    <span className="group-color" style={{ backgroundColor: colors.outflows }}></span>
                    <h4>Inside - Outflows</h4>
                    <span className="group-count">{results.outflows[0]?.cells.length || 0} cells</span>
                  </div>
                  <p className="group-description">Cells inside focus area that provide values to outside</p>
                </div>

                <div className="flow-group">
                  <div className="group-header">
                    <span className="group-color" style={{ backgroundColor: colors.outsideDependents }}></span>
                    <h4>Outside - Dependents</h4>
                    <span className="group-count">{results.outsideDependents?.[0]?.cells.length || 0} cells</span>
                  </div>
                  <p className="group-description">Cells outside focus area that are dependents of cells inside</p>
                </div>
              </div>
            )}

            <div className="color-actions">
              {!colorsApplied ? (
                <button className="btn btn-primary" onClick={this.applyColors}>
                  Apply Colors to Workbook
                </button>
              ) : (
                <>
                  <button className="btn btn-secondary" onClick={this.removeColors}>
                    Remove Colors
                  </button>
                  <button className="btn btn-primary" onClick={this.keepColors}>
                    Keep Colors (Alt+X)
                  </button>
                </>
              )}
            </div>
          </div>
        )}

        <div className="tool-help">
          <h4>How to Use</h4>
          <ul>
            <li>Select the scope of analysis (workbook, worksheet, or range)</li>
            <li>Optionally enable focus area to analyze inflows/outflows</li>
            <li>Click "Analyze Flow" to identify inputs, calculations, and outputs</li>
            <li>Apply colors to visualize the flow in your workbook</li>
            <li>Press <kbd>Alt+X</kbd> to keep colors permanently</li>
          </ul>
        </div>
      </div>
    );
  }
}
