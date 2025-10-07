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

  renderFormulaNodeRow = (node: FormulaNode, depth: number = 0, index: string = '0'): JSX.Element[] => {
    const { selectedNode } = this.state;
    const isSelected = selectedNode === node;
    const isActive = node.isActive;
    const rows: JSX.Element[] = [];

    // Render current node
    rows.push(
      <tr
        key={`row-${index}`}
        className={`formula-row ${isSelected ? 'selected' : ''} ${isActive ? 'active-branch' : ''}`}
        onClick={(e) => { e.stopPropagation(); this.selectNode(node, e); }}
      >
        {/* Element Column */}
        <td className="element-cell" style={{ paddingLeft: `${depth * 20 + 8}px` }}>
          <span className="tree-icon">
            {node.children && node.children.length > 0 ? '⊟' : ''}
          </span>
          <span className={`element-icon element-icon-${node.type}`}>
            {node.type === 'function' ? '⚡' : node.type === 'reference' ? '📍' : '•'}
          </span>
          <span className="element-value">{node.value}</span>
        </td>

        {/* Info Column */}
        <td className="info-cell">
          {node.argumentName || (node.type === 'operator' ? 'operator' : '')}
        </td>

        {/* Value Column */}
        <td className="value-cell">
          {node.calculatedValue !== undefined ? String(node.calculatedValue) : ''}
        </td>

        {/* Location Column */}
        <td className="location-cell">
          {node.location || (node.address ? node.address : '')}
        </td>
      </tr>
    );

    // Render children
    if (node.children && node.children.length > 0) {
      node.children.forEach((child, childIndex) => {
        const childRows = this.renderFormulaNodeRow(child, depth + 1, `${index}-${childIndex}`);
        rows.push(...childRows);
      });
    }

    return rows;
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
      <div className="formula-explorer-modern" onKeyDown={this.handleKeyDown}>
        {/* Formula Explorer Table - Main View */}
        <div className="explorer-table-container">
          <table className="explorer-table">
            <thead>
              <tr>
                <th className="col-element">Element</th>
                <th className="col-info">Info</th>
                <th className="col-value">Value</th>
                <th className="col-location">Location</th>
              </tr>
            </thead>
            <tbody>
              {formulaTree ? this.renderFormulaNodeRow(formulaTree) : (
                <tr><td colSpan={4} className="no-formula-row">No formula to display</td></tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Formula Display at Bottom */}
        <div className="formula-bottom-bar">
          {isEditing ? (
            <div className="formula-editor-inline">
              <textarea
                ref={this.formulaInputRef}
                className="formula-input-bottom"
                value={editedFormula}
                onChange={this.handleFormulaChange}
                rows={2}
              />
              <button className="btn-ok" onClick={this.applyEdit}>OK</button>
              <button className="btn-cancel" onClick={this.cancelEdit}>Cancel</button>
            </div>
          ) : (
            <div className="formula-display-bar">
              <code className="formula-code-bottom">{currentCell.formula || '(No formula)'}</code>
              <button className="btn-more" onClick={this.startEditing} title="Edit and expand formula">More ▼</button>
            </div>
          )}
        </div>
      </div>
    );
  }
}
