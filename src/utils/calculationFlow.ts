/**
 * Calculation Flow - Analyze data flow and dependencies in spreadsheets
 */

export interface FlowAnalysisResult {
  inputs: CellGroup[];
  calculations: CellGroup[];
  outputs: CellGroup[];
  orphans: CellGroup[];
  inflows?: CellGroup[];
  outflows?: CellGroup[];
  outsidePrecedents?: CellGroup[];
  outsideDependents?: CellGroup[];
}

export interface CellGroup {
  cells: string[];
  description: string;
  color: string;
}

export interface FlowColors {
  inputs: string;
  calculations: string;
  outputs: string;
  orphans: string;
  inflows: string;
  outflows: string;
  outsidePrecedents: string;
  outsideDependents: string;
}

export interface AnalysisScope {
  type: 'workbook' | 'worksheet' | 'range';
  sheetName?: string;
  rangeAddress?: string;
  focusArea?: {
    sheetName?: string;
    rangeAddress?: string;
  };
}

export class CalculationFlowAnalyzer {
  private static defaultColors: FlowColors = {
    inputs: '#C6E0B4',           // Light green
    calculations: '#FFE699',      // Light yellow
    outputs: '#F4B084',          // Light orange
    orphans: '#E7E6E6',          // Light gray
    inflows: '#BDD7EE',          // Light blue
    outflows: '#F8CBAD',         // Light peach
    outsidePrecedents: '#D5A6BD', // Light purple
    outsideDependents: '#B4C7E7'  // Light blue-gray
  };

  /**
   * Analyze calculation flow for a given scope
   */
  static async analyzeFlow(
    context: Excel.RequestContext,
    scope: AnalysisScope,
    colors: FlowColors = this.defaultColors
  ): Promise<FlowAnalysisResult> {
    const cellDependencies = await this.buildDependencyGraph(context, scope);
    
    const hasFocusArea = scope.focusArea !== undefined;

    if (hasFocusArea) {
      return this.analyzeInflowsOutflows(cellDependencies, scope, colors);
    } else {
      return this.analyzeInputsOutputs(cellDependencies, scope, colors);
    }
  }

  /**
   * Build dependency graph for the scope
   */
  private static async buildDependencyGraph(
    context: Excel.RequestContext,
    scope: AnalysisScope
  ): Promise<Map<string, { precedents: string[]; dependents: string[]; formula: string }>> {
    const dependencies = new Map<string, { precedents: string[]; dependents: string[]; formula: string }>();

    let sheets: string[] = [];

    if (scope.type === 'workbook') {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items/name');
      await context.sync();
      sheets = worksheets.items.map(ws => ws.name);
    } else {
      sheets = [scope.sheetName!];
    }

    for (const sheetName of sheets) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      let range: Excel.Range;

      if (scope.type === 'range' && scope.rangeAddress) {
        range = sheet.getRange(scope.rangeAddress);
      } else {
        range = sheet.getUsedRange();
      }

      range.load(['formulas', 'address', 'rowCount', 'columnCount']);
      await context.sync();

      const formulas = range.formulas as string[][];
      const baseAddress = range.address.split('!')[1].split(':')[0];
      const baseCoords = this.addressToCoords(baseAddress);

      for (let row = 0; row < range.rowCount; row++) {
        for (let col = 0; col < range.columnCount; col++) {
          const formula = formulas[row][col];
          
          if (!formula || !formula.startsWith('=')) {
            continue;
          }

          const cellAddress = this.getAddress(baseCoords.row + row, baseCoords.col + col);
          const fullAddress = `${sheetName}!${cellAddress}`;

          const precedents = this.extractPrecedents(formula, sheetName);
          
          dependencies.set(fullAddress, {
            precedents,
            dependents: [],
            formula
          });
        }
      }
    }

    // Build reverse dependencies (dependents)
    for (const [address, info] of dependencies.entries()) {
      for (const precedent of info.precedents) {
        const precedentInfo = dependencies.get(precedent);
        if (precedentInfo) {
          precedentInfo.dependents.push(address);
        } else {
          // Precedent is outside the scope
          dependencies.set(precedent, {
            precedents: [],
            dependents: [address],
            formula: ''
          });
        }
      }
    }

    return dependencies;
  }

  /**
   * Analyze inputs, calculations, outputs, and orphans
   */
  private static analyzeInputsOutputs(
    dependencies: Map<string, { precedents: string[]; dependents: string[]; formula: string }>,
    scope: AnalysisScope,
    colors: FlowColors
  ): FlowAnalysisResult {
    const inputs: string[] = [];
    const calculations: string[] = [];
    const outputs: string[] = [];
    const orphans: string[] = [];

    for (const [address, info] of dependencies.entries()) {
      if (!info.formula) {
        continue; // Skip cells outside scope
      }

      const inScopePrecedents = info.precedents.filter(p => 
        this.isInScope(p, scope) && dependencies.get(p)?.formula
      );
      const inScopeDependents = info.dependents.filter(d => 
        this.isInScope(d, scope) && dependencies.get(d)?.formula
      );

      const hasPrecedents = inScopePrecedents.length > 0;
      const hasDependents = inScopeDependents.length > 0;

      if (!hasPrecedents && !hasDependents) {
        orphans.push(address);
      } else if (!hasPrecedents && hasDependents) {
        inputs.push(address);
      } else if (hasPrecedents && !hasDependents) {
        outputs.push(address);
      } else {
        calculations.push(address);
      }
    }

    return {
      inputs: [{ cells: inputs, description: 'Inputs', color: colors.inputs }],
      calculations: [{ cells: calculations, description: 'Calculations', color: colors.calculations }],
      outputs: [{ cells: outputs, description: 'Outputs', color: colors.outputs }],
      orphans: [{ cells: orphans, description: 'Orphan Formulas', color: colors.orphans }]
    };
  }

  /**
   * Analyze inflows and outflows relative to focus area
   */
  private static analyzeInflowsOutflows(
    dependencies: Map<string, { precedents: string[]; dependents: string[]; formula: string }>,
    scope: AnalysisScope,
    colors: FlowColors
  ): FlowAnalysisResult {
    const inflows: string[] = [];
    const outflows: string[] = [];
    const outsidePrecedents: string[] = [];
    const outsideDependents: string[] = [];

    const focusArea = scope.focusArea!;

    for (const [address, info] of dependencies.entries()) {
      if (!info.formula) {
        continue;
      }

      const isInFocus = this.isInFocusArea(address, focusArea);

      if (isInFocus) {
        // Check for precedents outside focus area
        const outsidePrec = info.precedents.filter(p => 
          !this.isInFocusArea(p, focusArea) && this.isInScope(p, scope)
        );
        
        if (outsidePrec.length > 0) {
          inflows.push(address);
          outsidePrecedents.push(...outsidePrec);
        }

        // Check for dependents outside focus area
        const outsideDep = info.dependents.filter(d => 
          !this.isInFocusArea(d, focusArea) && this.isInScope(d, scope)
        );
        
        if (outsideDep.length > 0) {
          outflows.push(address);
          outsideDependents.push(...outsideDep);
        }
      }
    }

    // Also get the standard inputs/outputs analysis
    const standardAnalysis = this.analyzeInputsOutputs(dependencies, scope, colors);

    return {
      ...standardAnalysis,
      inflows: [{ cells: inflows, description: 'Inside - Inflows', color: colors.inflows }],
      outflows: [{ cells: outflows, description: 'Inside - Outflows', color: colors.outflows }],
      outsidePrecedents: [{ cells: [...new Set(outsidePrecedents)], description: 'Outside - Precedents', color: colors.outsidePrecedents }],
      outsideDependents: [{ cells: [...new Set(outsideDependents)], description: 'Outside - Dependents', color: colors.outsideDependents }]
    };
  }

  /**
   * Apply colors to cells based on analysis results
   */
  static async applyFlowColors(
    context: Excel.RequestContext,
    result: FlowAnalysisResult
  ): Promise<void> {
    const allGroups = [
      ...(result.inputs || []),
      ...(result.calculations || []),
      ...(result.outputs || []),
      ...(result.orphans || []),
      ...(result.inflows || []),
      ...(result.outflows || []),
      ...(result.outsidePrecedents || []),
      ...(result.outsideDependents || [])
    ];

    for (const group of allGroups) {
      for (const cellAddress of group.cells) {
        const [sheetName, address] = cellAddress.split('!');
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const range = sheet.getRange(address);
        range.format.fill.color = group.color;
      }
    }

    await context.sync();
  }

  /**
   * Remove flow colors from cells
   */
  static async removeFlowColors(
    context: Excel.RequestContext,
    scope: AnalysisScope
  ): Promise<void> {
    let sheets: string[] = [];

    if (scope.type === 'workbook') {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items/name');
      await context.sync();
      sheets = worksheets.items.map(ws => ws.name);
    } else {
      sheets = [scope.sheetName!];
    }

    for (const sheetName of sheets) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      let range: Excel.Range;

      if (scope.type === 'range' && scope.rangeAddress) {
        range = sheet.getRange(scope.rangeAddress);
      } else {
        range = sheet.getUsedRange();
      }

      range.format.fill.clear();
      await context.sync();
    }
  }

  /**
   * Extract precedents from a formula
   */
  private static extractPrecedents(formula: string, currentSheet: string): string[] {
    const precedents: string[] = [];
    const pattern = /(\[.*?\])?(\w+!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)/gi;
    
    let match;
    while ((match = pattern.exec(formula)) !== null) {
      const workbook = match[1] || '';
      const sheet = match[2] ? match[2].replace('!', '') : currentSheet;
      const address = match[3];

      // For ranges, we'll just use the whole range as one precedent
      precedents.push(`${sheet}!${address}`);
    }

    return precedents;
  }

  /**
   * Check if an address is in scope
   */
  private static isInScope(address: string, scope: AnalysisScope): boolean {
    const [sheetName, cellAddress] = address.split('!');

    if (scope.type === 'workbook') {
      return true;
    }

    if (scope.type === 'worksheet') {
      return sheetName === scope.sheetName;
    }

    if (scope.type === 'range') {
      return sheetName === scope.sheetName && 
             this.isAddressInRange(cellAddress, scope.rangeAddress!);
    }

    return false;
  }

  /**
   * Check if an address is in the focus area
   */
  private static isInFocusArea(address: string, focusArea: { sheetName?: string; rangeAddress?: string }): boolean {
    const [sheetName, cellAddress] = address.split('!');

    if (focusArea.sheetName && sheetName !== focusArea.sheetName) {
      return false;
    }

    if (focusArea.rangeAddress) {
      return this.isAddressInRange(cellAddress, focusArea.rangeAddress);
    }

    return true;
  }

  /**
   * Check if a cell address is within a range
   */
  private static isAddressInRange(cellAddress: string, rangeAddress: string): boolean {
    if (!rangeAddress.includes(':')) {
      return cellAddress === rangeAddress;
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
   * Convert cell address to coordinates
   */
  private static addressToCoords(address: string): { row: number; col: number } {
    const match = address.match(/([A-Z]+)(\d+)/i);
    if (!match) {
      throw new Error(`Invalid cell address: ${address}`);
    }

    const col = this.columnToNumber(match[1]);
    const row = parseInt(match[2], 10);

    return { row, col };
  }

  /**
   * Convert column letter to number
   */
  private static columnToNumber(column: string): number {
    let num = 0;
    for (let i = 0; i < column.length; i++) {
      num = num * 26 + (column.charCodeAt(i) - 64);
    }
    return num;
  }

  /**
   * Get cell address from coordinates
   */
  private static getAddress(row: number, col: number): string {
    const columnLetter = this.numberToColumn(col);
    return `${columnLetter}${row}`;
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
   * Get default colors
   */
  static getDefaultColors(): FlowColors {
    return { ...this.defaultColors };
  }
}
