/**
 * Excel Helper - Utilities for interacting with Excel API
 */

export interface CellInfo {
  address: string;
  formula: string;
  value: any;
  sheet: string;
  workbook: string;
}

export interface PrecedentInfo {
  address: string;
  sheet: string;
  workbook: string;
  value: any;
}

export class ExcelHelper {
  /**
   * Get information about the currently selected cell
   */
  static async getSelectedCellInfo(context: Excel.RequestContext): Promise<CellInfo> {
    const range = context.workbook.getSelectedRange();
    range.load(['address', 'formulas', 'values', 'worksheet']);
    
    await context.sync();

    const worksheet = range.worksheet;
    worksheet.load('name');
    
    await context.sync();

    // Extract just the cell address without sheet name
    const addressParts = range.address.split('!');
    const cellAddress = addressParts.length > 1 ? addressParts[1] : range.address;

    return {
      address: cellAddress,
      formula: range.formulas[0][0] as string,
      value: range.values[0][0],
      sheet: worksheet.name,
      workbook: 'Current Workbook'
    };
  }

  /**
   * Get direct precedents of a cell
   */
  static async getDirectPrecedents(
    context: Excel.RequestContext,
    address: string,
    sheetName?: string
  ): Promise<PrecedentInfo[]> {
    const precedents: PrecedentInfo[] = [];
    
    try {
      const sheet = sheetName 
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = sheet.getRange(address);
      range.load(['formulas', 'address']);
      
      await context.sync();

      const formula = range.formulas[0][0] as string;
      
      if (!formula || !formula.startsWith('=')) {
        return precedents;
      }

      // Extract cell references from formula
      const references = this.extractCellReferences(formula);
      
      for (const ref of references) {
        try {
          const { sheetName: refSheet, address: refAddress } = this.parseReference(ref, sheet.name);
          const refRange = refSheet 
            ? context.workbook.worksheets.getItem(refSheet).getRange(refAddress)
            : sheet.getRange(refAddress);
          
          refRange.load(['values', 'address']);
          await context.sync();

          precedents.push({
            address: refAddress,
            sheet: refSheet || sheet.name,
            workbook: 'Current Workbook',
            value: refRange.values[0][0]
          });
        } catch (error) {
          console.error(`Error loading precedent ${ref}:`, error);
        }
      }
    } catch (error) {
      console.error('Error getting precedents:', error);
    }

    return precedents;
  }

  /**
   * Get direct dependents of a cell or range
   */
  static async getDirectDependents(
    context: Excel.RequestContext,
    address: string,
    sheetName?: string
  ): Promise<CellInfo[]> {
    const dependents: CellInfo[] = [];
    
    try {
      const sheet = sheetName 
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const usedRange = sheet.getUsedRange();
      usedRange.load(['formulas', 'values', 'address', 'rowCount', 'columnCount']);
      
      await context.sync();

      const targetAddress = address.toUpperCase();
      
      // Scan all cells for references to the target
      for (let row = 0; row < usedRange.rowCount; row++) {
        for (let col = 0; col < usedRange.columnCount; col++) {
          const formula = usedRange.formulas[row][col] as string;
          
          if (formula && formula.startsWith('=')) {
            const references = this.extractCellReferences(formula);
            
            for (const ref of references) {
              const { address: refAddress } = this.parseReference(ref, sheet.name);
              
              if (refAddress.toUpperCase() === targetAddress || 
                  this.isInRange(targetAddress, refAddress)) {
                const cellAddress = this.getCellAddress(usedRange.address, row, col);
                
                dependents.push({
                  address: cellAddress,
                  formula: formula,
                  value: usedRange.values[row][col],
                  sheet: sheet.name,
                  workbook: 'Current Workbook'
                });
                break;
              }
            }
          }
        }
      }
    } catch (error) {
      console.error('Error getting dependents:', error);
    }

    return dependents;
  }

  /**
   * Extract cell references from a formula
   */
  static extractCellReferences(formula: string): string[] {
    const references: string[] = [];
    
    // Pattern to match cell references: A1, $A$1, Sheet1!A1, [Book1]Sheet1!A1, A1:B10
    const pattern = /(\[.*?\])?(\w+!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+:\$?[A-Z]+|\$?\d+:\$?\d+)/gi;
    
    let match;
    while ((match = pattern.exec(formula)) !== null) {
      references.push(match[0]);
    }

    return references;
  }

  /**
   * Parse a cell reference into sheet name and address
   */
  static parseReference(reference: string, defaultSheet: string): { sheetName: string; address: string } {
    const parts = reference.split('!');
    
    if (parts.length === 2) {
      let sheetName = parts[0];
      // Remove workbook reference if present
      if (sheetName.includes(']')) {
        sheetName = sheetName.substring(sheetName.indexOf(']') + 1);
      }
      // Remove quotes if present
      sheetName = sheetName.replace(/^'|'$/g, '');
      
      return {
        sheetName: sheetName,
        address: parts[1]
      };
    }

    return {
      sheetName: defaultSheet,
      address: reference
    };
  }

  /**
   * Check if an address is within a range
   */
  static isInRange(rangeAddress: string, cellAddress: string): boolean {
    if (!rangeAddress.includes(':')) {
      return rangeAddress.toUpperCase() === cellAddress.toUpperCase();
    }

    const [start, end] = rangeAddress.split(':');
    const startCoords = this.addressToCoords(start);
    const endCoords = this.addressToCoords(end);
    const cellCoords = this.addressToCoords(cellAddress);

    return cellCoords.row >= startCoords.row &&
           cellCoords.row <= endCoords.row &&
           cellCoords.col >= startCoords.col &&
           cellCoords.col <= endCoords.col;
  }

  /**
   * Convert cell address to row/column coordinates
   */
  static addressToCoords(address: string): { row: number; col: number } {
    const match = address.match(/([A-Z]+)(\d+)/i);
    if (!match) {
      throw new Error(`Invalid cell address: ${address}`);
    }

    const col = this.columnToNumber(match[1]);
    const row = parseInt(match[2], 10);

    return { row, col };
  }

  /**
   * Convert column letter to number (A=1, B=2, etc.)
   */
  static columnToNumber(column: string): number {
    let num = 0;
    for (let i = 0; i < column.length; i++) {
      num = num * 26 + (column.charCodeAt(i) - 64);
    }
    return num;
  }

  /**
   * Convert column number to letter (1=A, 2=B, etc.)
   */
  static numberToColumn(num: number): string {
    let column = '';
    while (num > 0) {
      const remainder = (num - 1) % 26;
      column = String.fromCharCode(65 + remainder) + column;
      num = Math.floor((num - 1) / 26);
    }
    return column;
  }

  /**
   * Get cell address from range and row/col offset
   */
  static getCellAddress(rangeAddress: string, rowOffset: number, colOffset: number): string {
    const baseAddress = rangeAddress.split(':')[0];
    const coords = this.addressToCoords(baseAddress);
    
    const newRow = coords.row + rowOffset;
    const newCol = coords.col + colOffset;
    
    return `${this.numberToColumn(newCol)}${newRow}`;
  }

  /**
   * Navigate to a specific cell
   */
  static async navigateToCell(
    context: Excel.RequestContext,
    address: string,
    sheetName?: string
  ): Promise<void> {
    try {
      const sheet = sheetName 
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = sheet.getRange(address);
      range.select();
      
      await context.sync();
    } catch (error) {
      console.error('Error navigating to cell:', error);
    }
  }

  /**
   * Highlight a range with a specific color
   */
  static async highlightRange(
    context: Excel.RequestContext,
    address: string,
    color: string,
    sheetName?: string
  ): Promise<void> {
    try {
      const sheet = sheetName 
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = sheet.getRange(address);
      range.format.fill.color = color;
      
      await context.sync();
    } catch (error) {
      console.error('Error highlighting range:', error);
    }
  }

  /**
   * Get value of a cell or range
   */
  static async getCellValue(
    context: Excel.RequestContext,
    address: string,
    sheetName?: string
  ): Promise<any> {
    try {
      const sheet = sheetName 
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = sheet.getRange(address);
      range.load('values');
      
      await context.sync();

      return range.values[0][0];
    } catch (error) {
      console.error('Error getting cell value:', error);
      return null;
    }
  }

  /**
   * Update cell formula
   */
  static async updateCellFormula(
    context: Excel.RequestContext,
    address: string,
    formula: string,
    sheetName?: string
  ): Promise<void> {
    try {
      const sheet = sheetName 
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = sheet.getRange(address);
      range.formulas = [[formula]];
      
      await context.sync();
    } catch (error) {
      console.error('Error updating cell formula:', error);
    }
  }

  /**
   * Check if a formula contains filtering functions
   */
  static isFilteringFunction(formula: string): boolean {
    const filteringFunctions = [
      'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS',
      'MAXIFS', 'MINIFS', 'FILTER', 'SUMPRODUCT'
    ];
    
    const upperFormula = formula.toUpperCase();
    return filteringFunctions.some(func => upperFormula.includes(func + '('));
  }

  /**
   * Identify cells meeting criteria for filtering functions
   */
  static async identifyFilteredCells(
    context: Excel.RequestContext,
    formula: string,
    sheetName?: string
  ): Promise<string[]> {
    // This is a simplified implementation
    // A full implementation would need to parse and evaluate the criteria
    const filteredCells: string[] = [];
    
    // Extract ranges from the formula
    const references = this.extractCellReferences(formula);
    
    for (const ref of references) {
      filteredCells.push(ref);
    }

    return filteredCells;
  }
}
