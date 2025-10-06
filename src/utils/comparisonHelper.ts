/**
 * Comparison Helper - Utilities for comparing workbooks, worksheets, and ranges
 */

export interface ComparisonResult {
  differences: DifferenceBlock[];
  totalDifferences: number;
  alignmentNeeded: boolean;
  insertedRows: number[];
  insertedColumns: number[];
}

export interface DifferenceBlock {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  referenceCells: CellDifference[];
  comparatorCells: CellDifference[];
}

export interface CellDifference {
  address: string;
  formula: string;
  value: any;
  isDifferent: boolean;
}

export interface ComparisonOptions {
  compareFormulas: boolean; // true = compare formulas, false = compare values
  ignoreInputs: boolean; // when comparing formulas, ignore differences in cell references
  detectAlignment: boolean; // detect inserted rows/columns
}

export class ComparisonHelper {
  /**
   * Compare two worksheets
   */
  static async compareWorksheets(
    context: Excel.RequestContext,
    referenceSheet: string,
    comparatorSheet: string,
    options: ComparisonOptions
  ): Promise<ComparisonResult> {
    const refSheet = context.workbook.worksheets.getItem(referenceSheet);
    const compSheet = context.workbook.worksheets.getItem(comparatorSheet);

    const refRange = refSheet.getUsedRange();
    const compRange = compSheet.getUsedRange();

    refRange.load(['formulas', 'values', 'address', 'rowCount', 'columnCount']);
    compRange.load(['formulas', 'values', 'address', 'rowCount', 'columnCount']);

    await context.sync();

    return this.compareRanges(
      refRange.formulas as string[][],
      refRange.values as any[][],
      compRange.formulas as string[][],
      compRange.values as any[][],
      options
    );
  }

  /**
   * Compare two ranges
   */
  static compareRanges(
    refFormulas: string[][],
    refValues: any[][],
    compFormulas: string[][],
    compValues: any[][],
    options: ComparisonOptions
  ): ComparisonResult {
    const differences: DifferenceBlock[] = [];
    let totalDifferences = 0;

    const maxRows = Math.max(refFormulas.length, compFormulas.length);
    const maxCols = Math.max(
      refFormulas[0]?.length || 0,
      compFormulas[0]?.length || 0
    );

    let currentBlock: DifferenceBlock | null = null;

    for (let row = 0; row < maxRows; row++) {
      for (let col = 0; col < maxCols; col++) {
        const refFormula = refFormulas[row]?.[col] || '';
        const refValue = refValues[row]?.[col];
        const compFormula = compFormulas[row]?.[col] || '';
        const compValue = compValues[row]?.[col];

        const isDifferent = options.compareFormulas
          ? this.areFormulasDifferent(refFormula, compFormula, options.ignoreInputs)
          : this.areValuesDifferent(refValue, compValue);

        if (isDifferent) {
          totalDifferences++;

          if (!currentBlock || !this.isAdjacent(currentBlock, row, col)) {
            if (currentBlock) {
              differences.push(currentBlock);
            }

            currentBlock = {
              startRow: row,
              startCol: col,
              endRow: row,
              endCol: col,
              referenceCells: [],
              comparatorCells: []
            };
          }

          currentBlock.endRow = Math.max(currentBlock.endRow, row);
          currentBlock.endCol = Math.max(currentBlock.endCol, col);

          currentBlock.referenceCells.push({
            address: this.getAddress(row, col),
            formula: refFormula,
            value: refValue,
            isDifferent: true
          });

          currentBlock.comparatorCells.push({
            address: this.getAddress(row, col),
            formula: compFormula,
            value: compValue,
            isDifferent: true
          });
        }
      }
    }

    if (currentBlock) {
      differences.push(currentBlock);
    }

    return {
      differences,
      totalDifferences,
      alignmentNeeded: false, // Would need more sophisticated detection
      insertedRows: [],
      insertedColumns: []
    };
  }

  /**
   * Compare two formulas
   */
  private static areFormulasDifferent(
    formula1: string,
    formula2: string,
    ignoreInputs: boolean
  ): boolean {
    if (!ignoreInputs) {
      return formula1 !== formula2;
    }

    // Normalize formulas by replacing cell references with placeholders
    const normalized1 = this.normalizeFormula(formula1);
    const normalized2 = this.normalizeFormula(formula2);

    return normalized1 !== normalized2;
  }

  /**
   * Normalize a formula by replacing cell references with placeholders
   */
  private static normalizeFormula(formula: string): string {
    if (!formula || !formula.startsWith('=')) {
      return formula;
    }

    // Replace cell references with REF
    const pattern = /(\[.*?\])?(\w+!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)/gi;
    return formula.replace(pattern, 'REF');
  }

  /**
   * Compare two values
   */
  private static areValuesDifferent(value1: any, value2: any): boolean {
    if (value1 === value2) {
      return false;
    }

    // Handle numeric comparison with tolerance
    if (typeof value1 === 'number' && typeof value2 === 'number') {
      return Math.abs(value1 - value2) > 1e-10;
    }

    // Handle null/undefined
    if ((value1 == null && value2 != null) || (value1 != null && value2 == null)) {
      return true;
    }

    return String(value1) !== String(value2);
  }

  /**
   * Check if a cell is adjacent to the current block
   */
  private static isAdjacent(block: DifferenceBlock, row: number, col: number): boolean {
    return (
      row >= block.startRow - 1 &&
      row <= block.endRow + 1 &&
      col >= block.startCol - 1 &&
      col <= block.endCol + 1
    );
  }

  /**
   * Get cell address from row and column indices
   */
  private static getAddress(row: number, col: number): string {
    const columnLetter = this.numberToColumn(col + 1);
    return `${columnLetter}${row + 1}`;
  }

  /**
   * Convert column number to letter
   */
  private static numberToColumn(num: number): string {
    let column = '';
    while (num > 0) {
      const remainder = (num - 1) % 26;
      column = String.fromCharCode(65 + remainder) + column;
      num = Math.floor((num - 1) / 26);
    }
    return column;
  }

  /**
   * Detect alignment issues (inserted/deleted rows and columns)
   */
  static detectAlignmentIssues(
    refFormulas: string[][],
    compFormulas: string[][]
  ): { insertedRows: number[]; insertedColumns: number[] } {
    const insertedRows: number[] = [];
    const insertedColumns: number[] = [];

    // This is a simplified implementation
    // A full implementation would use sequence alignment algorithms (e.g., Smith-Waterman)
    
    // Check for row differences
    if (refFormulas.length !== compFormulas.length) {
      // Detect which rows are different
      for (let i = 0; i < Math.max(refFormulas.length, compFormulas.length); i++) {
        const refRow = refFormulas[i];
        const compRow = compFormulas[i];

        if (!refRow || !compRow || !this.areRowsSimilar(refRow, compRow)) {
          if (compFormulas.length > refFormulas.length) {
            insertedRows.push(i);
          }
        }
      }
    }

    return { insertedRows, insertedColumns };
  }

  /**
   * Check if two rows are similar
   */
  private static areRowsSimilar(row1: string[], row2: string[]): boolean {
    if (row1.length !== row2.length) {
      return false;
    }

    let similarCount = 0;
    for (let i = 0; i < row1.length; i++) {
      if (this.normalizeFormula(row1[i]) === this.normalizeFormula(row2[i])) {
        similarCount++;
      }
    }

    // Consider rows similar if at least 70% of cells match
    return similarCount / row1.length >= 0.7;
  }

  /**
   * Insert alignment rows in a worksheet
   */
  static async insertAlignmentRows(
    context: Excel.RequestContext,
    sheetName: string,
    rowIndices: number[]
  ): Promise<void> {
    const sheet = context.workbook.worksheets.getItem(sheetName);

    for (const rowIndex of rowIndices.sort((a, b) => b - a)) {
      const range = sheet.getRangeByIndexes(rowIndex, 0, 1, 1);
      range.insert(Excel.InsertShiftDirection.down);
    }

    await context.sync();
  }

  /**
   * Remove alignment rows from a worksheet
   */
  static async removeAlignmentRows(
    context: Excel.RequestContext,
    sheetName: string
  ): Promise<void> {
    // Alignment rows would need to be marked somehow (e.g., with a specific color or tag)
    // This is a placeholder implementation
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();
    usedRange.load(['rowCount', 'format/fill/color']);

    await context.sync();

    // Remove rows with a specific marker color
    const markerColor = '#FFFF00'; // Yellow
    for (let i = usedRange.rowCount - 1; i >= 0; i--) {
      const row = sheet.getRangeByIndexes(i, 0, 1, 1);
      row.load('format/fill/color');
      await context.sync();

      if (row.format.fill.color === markerColor) {
        row.delete(Excel.DeleteShiftDirection.up);
        await context.sync();
      }
    }
  }

  /**
   * Copy formula from one side to another
   */
  static async copyFormula(
    context: Excel.RequestContext,
    fromSheet: string,
    toSheet: string,
    address: string
  ): Promise<void> {
    const sourceSheet = context.workbook.worksheets.getItem(fromSheet);
    const targetSheet = context.workbook.worksheets.getItem(toSheet);

    const sourceRange = sourceSheet.getRange(address);
    const targetRange = targetSheet.getRange(address);

    sourceRange.load('formulas');
    await context.sync();

    targetRange.formulas = sourceRange.formulas;
    await context.sync();
  }
}
