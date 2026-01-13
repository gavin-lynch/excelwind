import * as ExcelJS from 'exceljs';
import {
  AnyNode,
  CellNode,
  ColumnNode,
  GroupNode,
  ImageNode,
  RowNode,
  WorkbookNode,
  WorksheetNode,
} from './types';

export function getChildren(node: WorksheetNode | WorkbookNode | GroupNode | RowNode) {
  if (!node.props?.children) return undefined;

  if (Array.isArray(node.props.children)) {
    return node.props.children;
  }
  return [node.props.children];
}

export function getFormula(node: ColumnNode | RowNode | CellNode) {
  if (node.props.formula) {
    return node.props.formula;
  }
  return null;
}

/**
 *  @TODO implemeent
export function setFormula(cell: ExcelJS.Cell, formula: string) {
  cell.formula = formula;
}
 */

export function getFormat(node: ColumnNode | RowNode | CellNode) {
  if (node.props.format) {
    return node.props.format;
  }
  return null;
}

/**
 *  @TODO implemeent
 */
export function setWidth(data: Pick<ExcelJS.Column, 'width'>, width: number) {
  data.width = width;
}

/**
 *  @TODO implemeent
 */
export function setFormat(
  data: Pick<ExcelJS.Cell, 'numFmt'> | Pick<ExcelJS.Column, 'numFmt'>,
  format: string,
) {
  data.numFmt = format;
}

/**
 * @TODO implemeent
 */
export function getStyle(node: ColumnNode | RowNode | CellNode) {
  if (node.props.style) {
    return node.props.style;
  }
}

/**
 *  @TODO implemeent
 */
export function setStyle(
  data: Pick<ExcelJS.Cell, 'style'> | Pick<ExcelJS.Column, 'style'>,
  style: Partial<ExcelJS.Style>,
) {
  // biome-ignore lint/correctness/noUnusedVariables: Property 'numFmt' is explicitly omitted from the spread...
  const { numFmt, ...styleWithoutNumFmt } = style;
  data.style = styleWithoutNumFmt;
}

/**
 *  @TODO implemeent
 */
export function setValue(data: ExcelJS.Cell, value: CellNode['props']['value']) {
  data.value = value;
}

export function isImage(node: AnyNode): node is ImageNode {
  return node.type === 'Image';
}

export function isPrimitive(value: any): boolean {
  return (
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean' ||
    value instanceof Date
  );
}

export function isWorkbook(node: AnyNode): node is WorkbookNode {
  return node.type === 'Workbook';
}

export function isWorksheet(node: AnyNode): node is WorksheetNode {
  return node.type === 'Worksheet';
}

export function isGroup(node: AnyNode): node is GroupNode {
  return node.type === 'Group';
}

export function isRow(node: AnyNode): node is RowNode {
  return node.type === 'Row';
}

export function isCell(node: AnyNode): node is CellNode {
  return node.type === 'Cell';
}

export function isColumn(node: AnyNode): node is ColumnNode {
  return node.type === 'Column';
}

export function isObject(item: any): item is Record<string, any> {
  return item && typeof item === 'object' && !Array.isArray(item);
}

export function mergeDeep<T extends object = object>(...sources: any[]): T {
  const result: any = {};

  for (const source of sources) {
    if (isObject(source)) {
      for (const key in source) {
        const sourceValue = source[key];
        const resultValue = result[key];

        if (isObject(sourceValue) && isObject(resultValue)) {
          result[key] = mergeDeep(resultValue, sourceValue);
        } else {
          result[key] = sourceValue;
        }
      }
    }
  }
  return result as T;
}
