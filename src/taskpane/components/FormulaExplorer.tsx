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
  }

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
        <div className="explorer-header">
          <h2>Formula Explorer</h2>
          <div className="cell-info">
            <span className="cell-address">{currentCell.sheet}!{currentCell.address}</span>
            <span className="cell-value">Value: {String(currentCell.value)}</span>
          </div>
        </div>

        <div className="formula-tree-container">
          <h3>Formula Structure</h3>
          <div className="formula-tree">
            {formulaTree ? this.renderFormulaNode(formulaTree) : <p>No formula to display</p>}
          </div>
        </div>

        <div className="precedents-container">
          <h3>Precedents ({selectedPrecedents.length})</h3>
          <div className="precedents-list">
            {selectedPrecedents.map((precedent, index) => (
              <div key={index} className="precedent-item">
                <span className="precedent-address">
                  {precedent.sheet !== currentCell.sheet && `${precedent.sheet}!`}
                  {precedent.address}
                </span>
                <span className="precedent-value">= {String(precedent.value)}</span>
              </div>
            ))}
          </div>
        </div>

        <div className="formula-editor-container">
          <h3>Formula</h3>
          {isEditing ? (
            <div className="formula-editor">
              <textarea
                ref={this.formulaInputRef}
                className="formula-input"
                value={editedFormula}
                onChange={this.handleFormulaChange}
                rows={3}
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
            <div className="formula-display">
              <code className="formula-code">{currentCell.formula}</code>
              <button className="btn btn-secondary" onClick={this.startEditing}>
                Edit (F2)
              </button>
            </div>
          )}
        </div>

        <div className="explorer-actions">
          <button className="btn btn-primary" onClick={this.loadCurrentCell}>
            Refresh (Ctrl+Q)
          </button>
          <button
            className="btn btn-secondary"
            onClick={async () => {
              await Excel.run(async (context) => {
                await KeyboardShortcutManager.navigateBack(context);
              });
            }}
          >
            Back (Ctrl+Backspace)
          </button>
        </div>

        <div className="explorer-help">
          <h4>Keyboard Shortcuts</h4>
          <ul>
            <li><kbd>Ctrl+Q</kbd> - Open Formula Explorer</li>
            <li><kbd>F2</kbd> - Edit formula</li>
            <li><kbd>Enter</kbd> - Apply changes</li>
            <li><kbd>Esc</kbd> - Cancel / Close</li>
            <li><kbd>Ctrl+Backspace</kbd> - Navigate back</li>
          </ul>
        </div>
      </div>
    );
  }
}
