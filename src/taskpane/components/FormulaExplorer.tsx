import * as React from 'react';
import { FormulaParser, FormulaNode } from '../../utils/formulaParser';
import { ExcelHelper, CellInfo, PrecedentInfo } from '../../utils/excelHelper';
import { KeyboardShortcutManager } from '../../utils/keyboardShortcuts';

/* global Excel */

export interface FormulaExplorerState {
  currentCell: CellInfo | null;
  formulaTree: FormulaNode | null;
  selectedNode: FormulaNode | null;
  selectedPrecedents: PrecedentInfo[];
  isEditing: boolean;
  editedFormula: string;
  explorerWindows: number;
}

export class FormulaExplorer extends React.Component<{}, FormulaExplorerState> {
  private formulaInputRef: React.RefObject<HTMLTextAreaElement>;

  constructor(props: {}) {
    super(props);

    this.state = {
      currentCell: null,
      formulaTree: null,
      selectedNode: null,
      selectedPrecedents: [],
      isEditing: false,
      editedFormula: '',
      explorerWindows: 0
    };

    this.formulaInputRef = React.createRef();
  }

  componentDidMount() {
    this.loadCurrentCell();
    this.setupSelectionChangeListener();
  }

  componentWillUnmount() {
    // Clean up event handler
    if (this.selectionChangeHandler) {
      this.selectionChangeHandler.remove();
    }
  }

  private selectionChangeHandler: any = null;

  setupSelectionChangeListener = async () => {
    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        this.selectionChangeHandler = worksheet.onSelectionChanged.add(async () => {
          await this.loadCurrentCell();
        });
        await context.sync();
      });
    } catch (error) {
      console.error('Error setting up selection listener:', error);
    }
  };

  loadCurrentCell = async () => {
    try {
      await Excel.run(async (context) => {
        const cellInfo = await ExcelHelper.getSelectedCellInfo(context);
        
        // Add to navigation history
        KeyboardShortcutManager.addToHistory(cellInfo.sheet, cellInfo.address);

        // Parse formula
        const formulaTree = FormulaParser.parse(cellInfo.formula);

        // Load precedents
        const precedents = await ExcelHelper.getDirectPrecedents(
          context,
          cellInfo.address,
          cellInfo.sheet
        );

        this.setState({
          currentCell: cellInfo,
          formulaTree,
          selectedNode: formulaTree,
          selectedPrecedents: precedents,
          editedFormula: cellInfo.formula
        });
      });
    } catch (error) {
      console.error('Error loading cell:', error);
    }
  };

  selectNode = async (node: FormulaNode, event?: React.MouseEvent) => {
    if (event) {
      event.stopPropagation();
    }

    this.setState({ selectedNode: node });

    // If it's a reference, navigate to it
    if (node.type === 'reference' && node.address) {
      try {
        await Excel.run(async (context) => {
          const currentCell = this.state.currentCell;
          if (currentCell) {
            await ExcelHelper.navigateToCell(context, node.address!, currentCell.sheet);
            
            // Highlight the precedent
            await ExcelHelper.highlightRange(context, node.address!, '#ADD8E6', currentCell.sheet);
          }
        });
      } catch (error) {
        console.error('Error navigating to precedent:', error);
      }
    }
  };

  startEditing = () => {
    this.setState({ isEditing: true }, () => {
      if (this.formulaInputRef.current) {
        this.formulaInputRef.current.focus();
        this.formulaInputRef.current.select();
      }
    });
  };

  handleFormulaChange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    this.setState({ editedFormula: event.target.value });
  };

  applyEdit = async () => {
    const { currentCell, editedFormula } = this.state;
    
    if (!currentCell) return;

    try {
      await Excel.run(async (context) => {
        await ExcelHelper.updateCellFormula(
          context,
          currentCell.address,
          editedFormula,
          currentCell.sheet
        );

        // Reload the cell
        await this.loadCurrentCell();
        
        this.setState({ isEditing: false });
      });
    } catch (error) {
      console.error('Error updating formula:', error);
    }
  };

  cancelEdit = () => {
    this.setState({
      isEditing: false,
      editedFormula: this.state.currentCell?.formula || ''
    });
  };

  handleKeyDown = (event: React.KeyboardEvent) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      this.applyEdit();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.cancelEdit();
    } else if (event.key === 'F2') {
      event.preventDefault();
      this.startEditing();
    }
  };

  renderFormulaNode = (node: FormulaNode, depth: number = 0): JSX.Element => {
    const { selectedNode } = this.state;
    const isSelected = selectedNode === node;
    const isActive = node.isActive;

    return (
      <div
        key={`${node.value}-${depth}`}
        className={`formula-node ${isSelected ? 'selected' : ''} ${isActive ? 'active-branch' : ''}`}
        style={{ marginLeft: `${depth * 20}px` }}
        onClick={(e) => this.selectNode(node, e)}
      >
        <div className="node-header">
          <span className={`node-type node-type-${node.type}`}>
            {node.type === 'function' ? '𝑓' : node.type === 'reference' ? '📍' : '•'}
          </span>
          <span className="node-value">{node.value}</span>
          {node.calculatedValue !== undefined && (
            <span className="node-calculated-value">= {String(node.calculatedValue)}</span>
          )}
          {node.targetLocation && (
            <span className="node-target-location">→ {node.targetLocation}</span>
          )}
        </div>
        
        {node.children && node.children.length > 0 && (
          <div className="node-children">
            {node.children.map((child, index) => this.renderFormulaNode(child, depth + 1))}
          </div>
        )}
      </div>
    );
  };

  render() {
    const { currentCell, formulaTree, isEditing, editedFormula, selectedPrecedents } = this.state;

    if (!currentCell) {
      return (
        <div className="formula-explorer">
          <div className="empty-state">
            <p>Select a cell to explore its formula</p>
            <button className="btn btn-primary" onClick={this.loadCurrentCell}>
              Load Selected Cell
            </button>
          </div>
        </div>
      );
    }

    return (
      <div className="formula-explorer" onKeyDown={this.handleKeyDown}>
        {/* Cell Info Header */}
        <div className="cell-info-header">
          <div className="cell-address-badge">{currentCell.sheet}!{currentCell.address}</div>
          <div className="cell-result">
            <span className="result-label">Result:</span>
            <span className="result-value">{String(currentCell.value)}</span>
          </div>
        </div>

        {/* Formula Display - Prominent */}
        <div className="formula-section">
          <div className="formula-header">
            <h3>Formula</h3>
            {!isEditing && (
              <button className="btn-icon" onClick={this.startEditing} title="Edit Formula (F2)">
                ✏️
              </button>
            )}
          </div>
          {isEditing ? (
            <div className="formula-editor">
              <textarea
                ref={this.formulaInputRef}
                className="formula-input"
                value={editedFormula}
                onChange={this.handleFormulaChange}
                rows={4}
              />
              <div className="editor-actions">
                <button className="btn btn-primary" onClick={this.applyEdit}>
                  Apply (Enter)
                </button>
                <button className="btn btn-secondary" onClick={this.cancelEdit}>
                  Cancel (Esc)
                </button>
              </div>
            </div>
          ) : (
            <div className="formula-display-box">
              <code className="formula-code">{currentCell.formula || '(No formula)'}</code>
            </div>
          )}
        </div>

        {/* Cell References with Values */}
        {selectedPrecedents.length > 0 && (
          <div className="cell-references-section">
            <h3>Cell References ({selectedPrecedents.length})</h3>
            <div className="references-grid">
              {selectedPrecedents.map((precedent, index) => (
                <div key={index} className="reference-card">
                  <div className="reference-address">
                    {precedent.sheet !== currentCell.sheet && `${precedent.sheet}!`}
                    {precedent.address}
                  </div>
                  <div className="reference-value">{String(precedent.value)}</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Formula Structure Tree */}
        <div className="formula-structure-section">
          <h3>Formula Structure</h3>
          <div className="formula-tree">
            {formulaTree ? this.renderFormulaNode(formulaTree) : <p className="no-formula">No formula to display</p>}
          </div>
        </div>

        {/* Actions */}
        <div className="explorer-actions">
          <button className="btn btn-secondary" onClick={this.loadCurrentCell}>
            🔄 Refresh
          </button>
          <button
            className="btn btn-secondary"
            onClick={async () => {
              await Excel.run(async (context) => {
                await KeyboardShortcutManager.navigateBack(context);
              });
            }}
          >
            ← Back
          </button>
        </div>

        {/* Help Text */}
        <div className="explorer-tip">
          💡 Click on any cell in Excel to automatically update this view
        </div>
      </div>
    );
  }
}
