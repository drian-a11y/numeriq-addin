import * as React from 'react';
import { FormulaMapper, FormulaMapColors, FormulaCellInfo } from '../../utils/formulaMapper';

/* global Excel */

export interface FormulaMapToolState {
  selectedSheets: string[];
  availableSheets: string[];
  colors: FormulaMapColors;
  analyzeUnique: boolean;
  isMapping: boolean;
  isMapped: boolean;
  results: Map<string, FormulaCellInfo[]>;
  showColorPicker: boolean;
}

export class FormulaMapTool extends React.Component<{}, FormulaMapToolState> {
  constructor(props: {}) {
    super(props);

    this.state = {
      selectedSheets: [],
      availableSheets: [],
      colors: FormulaMapper.getDefaultColors(),
      analyzeUnique: true,
      isMapping: false,
      isMapped: false,
      results: new Map(),
      showColorPicker: false
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
          selectedSheets: [activeSheet.name]
        });
      });
    } catch (error) {
      console.error('Error loading sheets:', error);
    }
  };

  toggleSheetSelection = (sheetName: string) => {
    this.setState(prevState => {
      const isSelected = prevState.selectedSheets.includes(sheetName);
      return {
        selectedSheets: isSelected
          ? prevState.selectedSheets.filter(s => s !== sheetName)
          : [...prevState.selectedSheets, sheetName]
      };
    });
  };

  applyFormulaMap = async () => {
    const { selectedSheets, colors, analyzeUnique } = this.state;

    if (selectedSheets.length === 0) {
      alert('Please select at least one sheet');
      return;
    }

    this.setState({ isMapping: true });

    try {
      await Excel.run(async (context) => {
        const results = await FormulaMapper.applyFormulaMapToMultipleSheets(
          context,
          selectedSheets,
          colors,
          analyzeUnique
        );

        this.setState({
          results,
          isMapped: true,
          isMapping: false
        });
      });
    } catch (error) {
      console.error('Error applying formula map:', error);
      this.setState({ isMapping: false });
    }
  };

  removeFormulaMap = async () => {
    const { selectedSheets } = this.state;

    try {
      await Excel.run(async (context) => {
        for (const sheetName of selectedSheets) {
          await FormulaMapper.removeFormulaMap(context, sheetName);
        }

        this.setState({ isMapped: false, results: new Map() });
      });
    } catch (error) {
      console.error('Error removing formula map:', error);
    }
  };

  updateColor = (colorKey: keyof FormulaMapColors, value: string) => {
    this.setState(prevState => ({
      colors: {
        ...prevState.colors,
        [colorKey]: value
      }
    }));
  };

  resetColors = () => {
    this.setState({ colors: FormulaMapper.getDefaultColors() });
  };

  getStatistics = (): { total: number; unique: number; copied: number; external: number; noRefs: number } => {
    const { results } = this.state;
    let total = 0;
    let unique = 0;
    let copied = 0;
    let external = 0;
    let noRefs = 0;

    for (const cellInfos of results.values()) {
      total += cellInfos.length;
      
      const formulaGroups = new Map<string, number>();
      for (const cell of cellInfos) {
        const count = formulaGroups.get(cell.normalizedFormula) || 0;
        formulaGroups.set(cell.normalizedFormula, count + 1);
        
        if (cell.hasExternalRef) external++;
        if (cell.hasNoReferences) noRefs++;
      }

      for (const count of formulaGroups.values()) {
        if (count === 1) unique++;
        else copied += count;
      }
    }

    return { total, unique, copied, external, noRefs };
  };

  render() {
    const {
      selectedSheets,
      availableSheets,
      colors,
      analyzeUnique,
      isMapping,
      isMapped,
      results,
      showColorPicker
    } = this.state;

    const stats = results.size > 0 ? this.getStatistics() : null;

    return (
      <div className="formula-map-tool">
        <div className="tool-header">
          <h2>Formula Map</h2>
          <p className="tool-description">
            Apply color schemes to reveal formula patterns and inconsistencies
          </p>
        </div>

        <div className="sheet-selection">
          <h3>Select Worksheets</h3>
          <div className="sheet-list">
            {availableSheets.map(sheetName => (
              <label key={sheetName} className="checkbox-label sheet-checkbox">
                <input
                  type="checkbox"
                  checked={selectedSheets.includes(sheetName)}
                  onChange={() => this.toggleSheetSelection(sheetName)}
                />
                <span className="sheet-name">{sheetName}</span>
              </label>
            ))}
          </div>
        </div>

        <div className="mapping-options">
          <h3>Options</h3>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={analyzeUnique}
              onChange={(e) => this.setState({ analyzeUnique: e.target.checked })}
            />
            Analyze unique formulas (identify external references and formulas with no references)
          </label>
        </div>

        <div className="color-configuration">
          <div className="section-header">
            <h3>Color Scheme</h3>
            <button
              className="btn btn-link"
              onClick={() => this.setState({ showColorPicker: !showColorPicker })}
            >
              {showColorPicker ? 'Hide' : 'Customize'} Colors
            </button>
            <button className="btn btn-link" onClick={this.resetColors}>
              Reset to Default
            </button>
          </div>

          {showColorPicker && (
            <div className="color-picker-grid">
              <div className="color-picker-item">
                <label>Unique Formula</label>
                <input
                  type="color"
                  value={colors.uniqueFormula}
                  onChange={(e) => this.updateColor('uniqueFormula', e.target.value)}
                />
                <input
                  type="text"
                  value={colors.uniqueFormula}
                  onChange={(e) => this.updateColor('uniqueFormula', e.target.value)}
                  className="color-input"
                />
              </div>

              <div className="color-picker-item">
                <label>Copied Formula</label>
                <input
                  type="color"
                  value={colors.copiedFormula}
                  onChange={(e) => this.updateColor('copiedFormula', e.target.value)}
                />
                <input
                  type="text"
                  value={colors.copiedFormula}
                  onChange={(e) => this.updateColor('copiedFormula', e.target.value)}
                  className="color-input"
                />
              </div>

              <div className="color-picker-item">
                <label>External Reference</label>
                <input
                  type="color"
                  value={colors.externalReference}
                  onChange={(e) => this.updateColor('externalReference', e.target.value)}
                />
                <input
                  type="text"
                  value={colors.externalReference}
                  onChange={(e) => this.updateColor('externalReference', e.target.value)}
                  className="color-input"
                />
              </div>

              <div className="color-picker-item">
                <label>No References</label>
                <input
                  type="color"
                  value={colors.noReferences}
                  onChange={(e) => this.updateColor('noReferences', e.target.value)}
                />
                <input
                  type="text"
                  value={colors.noReferences}
                  onChange={(e) => this.updateColor('noReferences', e.target.value)}
                  className="color-input"
                />
              </div>

              <div className="color-picker-item">
                <label>Hardcoded Value</label>
                <input
                  type="color"
                  value={colors.hardcodedValue}
                  onChange={(e) => this.updateColor('hardcodedValue', e.target.value)}
                />
                <input
                  type="text"
                  value={colors.hardcodedValue}
                  onChange={(e) => this.updateColor('hardcodedValue', e.target.value)}
                  className="color-input"
                />
              </div>
            </div>
          )}

          <div className="color-legend">
            <div className="legend-item">
              <span className="legend-color" style={{ backgroundColor: colors.uniqueFormula }}></span>
              <span className="legend-label">Unique Formula</span>
            </div>
            <div className="legend-item">
              <span className="legend-color" style={{ backgroundColor: colors.copiedFormula }}></span>
              <span className="legend-label">Copied Formula</span>
            </div>
            <div className="legend-item">
              <span className="legend-color" style={{ backgroundColor: colors.externalReference }}></span>
              <span className="legend-label">External Reference</span>
            </div>
            <div className="legend-item">
              <span className="legend-color" style={{ backgroundColor: colors.noReferences }}></span>
              <span className="legend-label">No References</span>
            </div>
            <div className="legend-item">
              <span className="legend-color" style={{ backgroundColor: colors.hardcodedValue }}></span>
              <span className="legend-label">Hardcoded Value</span>
            </div>
          </div>
        </div>

        <div className="mapping-actions">
          <button
            className="btn btn-primary btn-large"
            onClick={this.applyFormulaMap}
            disabled={selectedSheets.length === 0 || isMapping}
          >
            {isMapping ? 'Applying Map...' : isMapped ? 'Reapply Formula Map' : 'Apply Formula Map'}
          </button>
          {isMapped && (
            <button className="btn btn-secondary btn-large" onClick={this.removeFormulaMap}>
              Remove Formula Map
            </button>
          )}
        </div>

        {stats && (
          <div className="mapping-results">
            <h3>Statistics</h3>
            <div className="stats-grid">
              <div className="stat-item">
                <span className="stat-value">{stats.total}</span>
                <span className="stat-label">Total Formulas</span>
              </div>
              <div className="stat-item">
                <span className="stat-value">{stats.unique}</span>
                <span className="stat-label">Unique Formulas</span>
              </div>
              <div className="stat-item">
                <span className="stat-value">{stats.copied}</span>
                <span className="stat-label">Copied Formulas</span>
              </div>
              <div className="stat-item">
                <span className="stat-value">{stats.external}</span>
                <span className="stat-label">External References</span>
              </div>
              <div className="stat-item">
                <span className="stat-value">{stats.noRefs}</span>
                <span className="stat-label">No References</span>
              </div>
            </div>
          </div>
        )}

        <div className="tool-help">
          <h4>How to Use</h4>
          <ul>
            <li>Select one or more worksheets to analyze</li>
            <li>Customize colors if desired</li>
            <li>Click "Apply Formula Map" to color-code formulas</li>
            <li>Unique formulas are highlighted to reveal potential inconsistencies</li>
            <li>Press <kbd>Ctrl+Shift+M</kbd> to toggle the formula map</li>
          </ul>
        </div>
      </div>
    );
  }
}
