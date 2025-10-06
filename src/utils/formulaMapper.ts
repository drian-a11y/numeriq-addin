/**
 * Formula Mapper - Apply color schemes to reveal formula patterns
 */

export interface FormulaMapColors {
  uniqueFormula: string;
  copiedFormula: string;
  externalReference: string;
  noReferences: string;
  hardcodedValue: string;
  noFill: string;
}

export interface FormulaCellInfo {
  address: string;
  formula: string;
  normalizedFormula: string;
  hasExternalRef: boolean;
  hasNoReferences: boolean;
  isHardcoded: boolean;
  color: string;
}

export class FormulaMapper {
  private static defaultColors: FormulaMapColors = {
    uniqueFormula: '#FFE699',      // Light yellow
    copiedFormula: '#C6E0B4',      // Light green
    externalReference: '#F4B084',  // Light orange
    noReferences: '#BDD7EE',       // Light blue
    hardcodedValue: '#E7E6E6',     // Light gray
    noFill: ''
  };

  private static originalColors: Map<string, string> = new Map();

  /**
   * Apply formula mapping to a worksheet
   */
  static async applyFormulaMap(
    context: Excel.RequestContext,
    sheetName: string,
    colors: FormulaMapColors = this.defaultColors,
    analyzeUnique: boolean = true
  ): Promise<FormulaCellInfo[]> {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();

    usedRange.load(['formulas', 'values', 'address', 'rowCount', 'columnCount', 'format/fill/color']);
    await context.sync();

    // Store original colors
    this.storeOriginalColors(usedRange, sheetName);

    const formulas = usedRange.formulas as string[][];
    const values = usedRange.values as any[][];
    const cellInfos: FormulaCellInfo[] = [];

    // First pass: categorize all formulas
    const formulaMap = new Map<string, { count: number; cells: { row: number; col: number }[] }>();

    for (let row = 0; row < usedRange.rowCount; row++) {
      for (let col = 0; col < usedRange.columnCount; col++) {
        const formula = formulas[row][col];
        const value = values[row][col];

        if (!formula || !formula.startsWith('=')) {
          // Hardcoded value
          if (value !== null && value !== undefined && value !== '') {
            const cellRange = sheet.getRangeByIndexes(row, col, 1, 1);
            if (colors.hardcodedValue) {
              cellRange.format.fill.color = colors.hardcodedValue;
            }
          }
          continue;
        }

        const normalized = this.normalizeFormula(formula);
        
        if (!formulaMap.has(normalized)) {
          formulaMap.set(normalized, { count: 0, cells: [] });
        }

        const entry = formulaMap.get(normalized)!;
        entry.count++;
        entry.cells.push({ row, col });
      }
    }

    await context.sync();

    // Second pass: apply colors based on categorization
    for (const [normalized, entry] of formulaMap.entries()) {
      const isUnique = entry.count === 1;
      
      for (const { row, col } of entry.cells) {
        const formula = formulas[row][col];
        const address = this.getAddress(row, col);
        
        const hasExternalRef = this.hasExternalReference(formula);
        const hasNoReferences = !this.hasAnyReferences(formula);
        
        let color = '';
        
        if (analyzeUnique) {
          if (hasExternalRef && colors.externalReference) {
            color = colors.externalReference;
          } else if (hasNoReferences && colors.noReferences) {
            color = colors.noReferences;
          } else if (isUnique && colors.uniqueFormula) {
            color = colors.uniqueFormula;
          } else if (!isUnique && colors.copiedFormula) {
            color = colors.copiedFormula;
          }
        } else {
          if (isUnique && colors.uniqueFormula) {
            color = colors.uniqueFormula;
          } else if (!isUnique && colors.copiedFormula) {
            color = colors.copiedFormula;
          }
        }

        if (color) {
          const cellRange = sheet.getRangeByIndexes(row, col, 1, 1);
          cellRange.format.fill.color = color;
        }

        cellInfos.push({
          address,
          formula,
          normalizedFormula: normalized,
          hasExternalRef,
          hasNoReferences,
          isHardcoded: false,
          color
        });
      }
    }

    await context.sync();

    return cellInfos;
  }

  /**
   * Apply formula mapping to multiple worksheets
   */
  static async applyFormulaMapToMultipleSheets(
    context: Excel.RequestContext,
    sheetNames: string[],
    colors: FormulaMapColors = this.defaultColors,
    analyzeUnique: boolean = true
  ): Promise<Map<string, FormulaCellInfo[]>> {
    const results = new Map<string, FormulaCellInfo[]>();

    for (const sheetName of sheetNames) {
      const cellInfos = await this.applyFormulaMap(context, sheetName, colors, analyzeUnique);
      results.set(sheetName, cellInfos);
    }

    return results;
  }

  /**
   * Remove formula mapping and restore original colors
   */
  static async removeFormulaMap(
    context: Excel.RequestContext,
    sheetName: string
  ): Promise<void> {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();

    usedRange.load(['address', 'rowCount', 'columnCount']);
    await context.sync();

    // Restore original colors
    for (let row = 0; row < usedRange.rowCount; row++) {
      for (let col = 0; col < usedRange.columnCount; col++) {
        const address = this.getAddress(row, col);
        const key = `${sheetName}!${address}`;
        const originalColor = this.originalColors.get(key);

        if (originalColor !== undefined) {
          const cellRange = sheet.getRangeByIndexes(row, col, 1, 1);
          cellRange.format.fill.color = originalColor;
        }
      }
    }

    await context.sync();

    // Clear stored colors for this sheet
    for (const key of this.originalColors.keys()) {
      if (key.startsWith(`${sheetName}!`)) {
        this.originalColors.delete(key);
      }
    }
  }

  /**
   * Store original colors before applying mapping
   */
  private static storeOriginalColors(range: Excel.Range, sheetName: string): void {
    const rowCount = range.rowCount;
    const columnCount = range.columnCount;

    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < columnCount; col++) {
        const address = this.getAddress(row, col);
        const key = `${sheetName}!${address}`;
        
        // This would need to be loaded from the range
        // For now, we'll assume white/no fill as default
        if (!this.originalColors.has(key)) {
          this.originalColors.set(key, '#FFFFFF');
        }
      }
    }
  }

  /**
   * Normalize a formula for comparison
   */
  private static normalizeFormula(formula: string): string {
    if (!formula || !formula.startsWith('=')) {
      return formula;
    }

    // Replace cell references with placeholders while preserving structure
    let normalized = formula;

    // Replace absolute and relative references
    normalized = normalized.replace(/\$?[A-Z]+\$?\d+/g, 'REF');
    
    // Replace range references
    normalized = normalized.replace(/REF:REF/g, 'RANGE');

    return normalized;
  }

  /**
   * Check if formula has external references
   */
  private static hasExternalReference(formula: string): boolean {
    if (!formula || !formula.startsWith('=')) {
      return false;
    }

    // Check for workbook references: [Book1]Sheet1!A1
    if (formula.includes('[') && formula.includes(']')) {
      return true;
    }

    // Check for sheet references: Sheet1!A1
    if (formula.includes('!')) {
      return true;
    }

    return false;
  }

  /**
   * Check if formula has any cell references
   */
  private static hasAnyReferences(formula: string): boolean {
    if (!formula || !formula.startsWith('=')) {
      return false;
    }

    // Check for cell references
    const pattern = /[A-Z]+\d+/;
    return pattern.test(formula);
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
   * Get custom colors
   */
  static getDefaultColors(): FormulaMapColors {
    return { ...this.defaultColors };
  }

  /**
   * Set custom colors
   */
  static setCustomColors(colors: Partial<FormulaMapColors>): FormulaMapColors {
    return { ...this.defaultColors, ...colors };
  }
}
