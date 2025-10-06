/**
 * Formula Parser - Parses Excel formulas into a logical tree structure
 */

export interface FormulaNode {
  type: 'function' | 'reference' | 'operator' | 'literal' | 'array';
  value: string;
  children?: FormulaNode[];
  address?: string;
  calculatedValue?: any;
  isActive?: boolean; // For IF, IFS, CHOOSE, SWITCH - indicates which branch is active
  targetLocation?: string; // For VLOOKUP, OFFSET, INDEX, INDIRECT
}

export class FormulaParser {
  /**
   * Parse a formula string into a tree structure
   */
  static parse(formula: string): FormulaNode {
    if (!formula || !formula.startsWith('=')) {
      return {
        type: 'literal',
        value: formula
      };
    }

    const cleanFormula = formula.substring(1); // Remove leading =
    return this.parseExpression(cleanFormula);
  }

  private static parseExpression(expr: string): FormulaNode {
    expr = expr.trim();

    // Check if it's a function
    const funcMatch = expr.match(/^([A-Z_][A-Z0-9_.]*)\s*\(/i);
    if (funcMatch) {
      return this.parseFunction(expr);
    }

    // Check if it's a cell reference
    if (this.isCellReference(expr)) {
      return {
        type: 'reference',
        value: expr,
        address: expr
      };
    }

    // Check if it's a literal (number, string, boolean)
    if (this.isLiteral(expr)) {
      return {
        type: 'literal',
        value: expr
      };
    }

    // Check for operators
    const operatorMatch = this.findTopLevelOperator(expr);
    if (operatorMatch) {
      return {
        type: 'operator',
        value: operatorMatch.operator,
        children: [
          this.parseExpression(expr.substring(0, operatorMatch.index)),
          this.parseExpression(expr.substring(operatorMatch.index + operatorMatch.operator.length))
        ]
      };
    }

    // Default: treat as literal
    return {
      type: 'literal',
      value: expr
    };
  }

  private static parseFunction(expr: string): FormulaNode {
    const funcMatch = expr.match(/^([A-Z_][A-Z0-9_.]*)\s*\(/i);
    if (!funcMatch) {
      throw new Error('Invalid function expression');
    }

    const funcName = funcMatch[1].toUpperCase();
    const argsStart = funcMatch[0].length;
    const argsEnd = this.findMatchingParen(expr, argsStart - 1);
    const argsString = expr.substring(argsStart, argsEnd);
    
    const args = this.splitArguments(argsString);
    const children = args.map(arg => this.parseExpression(arg));

    const node: FormulaNode = {
      type: 'function',
      value: funcName,
      children
    };

    return node;
  }

  private static findMatchingParen(str: string, openIndex: number): number {
    let depth = 1;
    let inString = false;
    let stringChar = '';

    for (let i = openIndex + 1; i < str.length; i++) {
      const char = str[i];

      if (inString) {
        if (char === stringChar && str[i - 1] !== '\\') {
          inString = false;
        }
      } else {
        if (char === '"' || char === "'") {
          inString = true;
          stringChar = char;
        } else if (char === '(') {
          depth++;
        } else if (char === ')') {
          depth--;
          if (depth === 0) {
            return i;
          }
        }
      }
    }

    return str.length;
  }

  private static splitArguments(argsString: string): string[] {
    const args: string[] = [];
    let currentArg = '';
    let depth = 0;
    let inString = false;
    let stringChar = '';

    for (let i = 0; i < argsString.length; i++) {
      const char = argsString[i];

      if (inString) {
        currentArg += char;
        if (char === stringChar && argsString[i - 1] !== '\\') {
          inString = false;
        }
      } else {
        if (char === '"' || char === "'") {
          inString = true;
          stringChar = char;
          currentArg += char;
        } else if (char === '(' || char === '{') {
          depth++;
          currentArg += char;
        } else if (char === ')' || char === '}') {
          depth--;
          currentArg += char;
        } else if (char === ',' && depth === 0) {
          args.push(currentArg.trim());
          currentArg = '';
        } else {
          currentArg += char;
        }
      }
    }

    if (currentArg.trim()) {
      args.push(currentArg.trim());
    }

    return args;
  }

  private static findTopLevelOperator(expr: string): { operator: string; index: number } | null {
    const operators = ['+', '-', '*', '/', '^', '&', '=', '<>', '<=', '>=', '<', '>'];
    let depth = 0;
    let inString = false;
    let stringChar = '';

    // Search from right to left for lower precedence operators first
    const precedenceOrder = ['=', '<>', '<=', '>=', '<', '>', '&', '+', '-', '*', '/', '^'];

    for (const op of precedenceOrder) {
      depth = 0;
      inString = false;

      for (let i = expr.length - 1; i >= 0; i--) {
        const char = expr[i];

        if (inString) {
          if (char === stringChar && (i === 0 || expr[i - 1] !== '\\')) {
            inString = false;
          }
        } else {
          if (char === '"' || char === "'") {
            inString = true;
            stringChar = char;
          } else if (char === ')') {
            depth++;
          } else if (char === '(') {
            depth--;
          } else if (depth === 0) {
            if (op.length === 2 && i > 0 && expr.substring(i - 1, i + 1) === op) {
              return { operator: op, index: i - 1 };
            } else if (op.length === 1 && char === op) {
              return { operator: op, index: i };
            }
          }
        }
      }
    }

    return null;
  }

  private static isCellReference(expr: string): boolean {
    // Match patterns like A1, $A$1, Sheet1!A1, [Book1]Sheet1!A1, A:A, 1:1, A1:B10
    const cellRefPattern = /^(\[.*?\])?(\w+!)?(\$?[A-Z]+\$?\d+|\$?[A-Z]+:\$?[A-Z]+|\$?\d+:\$?\d+|\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i;
    return cellRefPattern.test(expr.trim());
  }

  private static isLiteral(expr: string): boolean {
    expr = expr.trim();
    
    // Check for number
    if (!isNaN(Number(expr))) {
      return true;
    }

    // Check for string (quoted)
    if ((expr.startsWith('"') && expr.endsWith('"')) || 
        (expr.startsWith("'") && expr.endsWith("'"))) {
      return true;
    }

    // Check for boolean
    if (expr.toUpperCase() === 'TRUE' || expr.toUpperCase() === 'FALSE') {
      return true;
    }

    return false;
  }

  /**
   * Evaluate which branch is active for logical functions
   */
  static evaluateLogicalBranch(funcName: string, args: any[], node: FormulaNode): void {
    switch (funcName) {
      case 'IF':
        if (args.length >= 2 && node.children) {
          const condition = args[0];
          node.children[1].isActive = !!condition;
          if (node.children[2]) {
            node.children[2].isActive = !condition;
          }
        }
        break;

      case 'IFS':
        if (node.children) {
          for (let i = 0; i < node.children.length; i += 2) {
            if (args[i]) {
              node.children[i].isActive = true;
              if (node.children[i + 1]) {
                node.children[i + 1].isActive = true;
              }
              break;
            }
          }
        }
        break;

      case 'CHOOSE':
        if (args.length >= 2 && node.children) {
          const index = Math.floor(args[0]);
          if (index >= 1 && index < node.children.length) {
            node.children[index].isActive = true;
          }
        }
        break;

      case 'SWITCH':
        if (args.length >= 3 && node.children) {
          const expr = args[0];
          for (let i = 1; i < args.length - 1; i += 2) {
            if (args[i] === expr) {
              node.children[i].isActive = true;
              node.children[i + 1].isActive = true;
              break;
            }
          }
          // Check for default value
          if (args.length % 2 === 0 && node.children[node.children.length - 1]) {
            node.children[node.children.length - 1].isActive = true;
          }
        }
        break;
    }
  }

  /**
   * Evaluate target location for reference functions
   */
  static evaluateTargetLocation(funcName: string, args: any[], node: FormulaNode): string | null {
    switch (funcName) {
      case 'VLOOKUP':
      case 'HLOOKUP':
        // args: lookup_value, table_array, col_index_num, [range_lookup]
        if (args.length >= 3 && node.children && node.children[1]) {
          const tableArray = node.children[1].value;
          const colIndex = args[2];
          return `${tableArray} (column ${colIndex})`;
        }
        break;

      case 'INDEX':
        // args: array, row_num, [column_num]
        if (args.length >= 2 && node.children && node.children[0]) {
          const array = node.children[0].value;
          const row = args[1];
          const col = args.length >= 3 ? args[2] : 1;
          return `${array} (row ${row}, col ${col})`;
        }
        break;

      case 'OFFSET':
        // args: reference, rows, cols, [height], [width]
        if (args.length >= 3 && node.children && node.children[0]) {
          const reference = node.children[0].value;
          const rows = args[1];
          const cols = args[2];
          return `${reference} offset by (${rows}, ${cols})`;
        }
        break;

      case 'INDIRECT':
        // args: ref_text, [a1]
        if (args.length >= 1) {
          return `Indirect: ${args[0]}`;
        }
        break;
    }

    return null;
  }
}
