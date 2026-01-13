import type { Alignment, Border, Font, Style } from 'exceljs';
import { formatHex } from 'culori';
import twColors from 'tailwindcss/colors';

// --- Color Map Generation ---
const generateColorMap = (): Record<string, string> => {
  const colorMap: Record<string, string> = {};

  for (const [colorName, colorValue] of Object.entries(twColors)) {
    if (typeof colorValue === 'string') {
      // Handles simple colors like 'black', 'white', and oklch strings
      try {
        const hex = formatHex(colorValue);
        if (hex) {
          colorMap[colorName] = hex.substring(1).toUpperCase();
        }
      } catch (_e) {
        // culori will throw on non-color strings like 'inherit', which we can ignore.
      }
    } else if (typeof colorValue === 'object' && colorValue !== null) {
      // Handle colors with shades like 'blue', 'red', etc.
      for (const [shade, oklch] of Object.entries(colorValue)) {
        if (typeof oklch === 'string') {
          try {
            const hex = formatHex(oklch);
            if (hex) {
              colorMap[`${colorName}-${shade}`] = hex.substring(1).toUpperCase();
            }
          } catch (_e) {
            // Ignore if culori fails to parse
          }
        }
      }
    }
  }

  return colorMap;
};

const colors = generateColorMap();

const fontSizes: Record<string, number> = {
  'text-xs': 10,
  'text-sm': 11,
  'text-base': 12,
  'text-lg': 14,
  'text-xl': 16,
  'text-2xl': 20,
  'text-3xl': 24,
  'text-4xl': 30,
};

const fontStyles: Record<string, Partial<Font>> = {
  'font-bold': { bold: true },
  'font-italic': { italic: true },
  'font-underline': { underline: true },
};

const alignmentStyles: Record<string, Partial<Alignment>> = {
  'text-left': { horizontal: 'left' },
  'text-center': { horizontal: 'center' },
  'text-right': { horizontal: 'right' },
  'align-top': { vertical: 'top' },
  'align-middle': { vertical: 'middle' },
  'align-center': { vertical: 'middle' },
  'align-bottom': { vertical: 'bottom' },
  'text-wrap': { wrapText: true },
  'text-nowrap': { wrapText: false },
};

const borderStyles: Record<string, Border['style']> = {
  solid: 'thin',
  thin: 'thin',
  dotted: 'dotted',
  dashed: 'dashed',
  double: 'double',
  thick: 'thick',
  hair: 'hair',
  medium: 'medium',
  'dash-dot': 'dashDot',
  'medium-dashed': 'mediumDashed',
  'dash-dot-dot': 'dashDotDot',
  'medium-dash-dot': 'mediumDashDot',
  'slant-dash-dot': 'slantDashDot',
  'medium-dash-dot-dot': 'mediumDashDotDot',
};

type BorderSide = 'top' | 'left' | 'right' | 'bottom';

const borderSideMap: Record<string, BorderSide[]> = {
  border: ['top', 'left', 'right', 'bottom'],
  'border-t': ['top'],
  'border-b': ['bottom'],
  'border-l': ['left'],
  'border-r': ['right'],
  'border-x': ['left', 'right'],
  'border-y': ['top', 'bottom'],
};

// --- Main Function ---
export function tailwindExcel(classString: string): Partial<Style> {
  const classes = classString.split(' ').filter(Boolean);
  const style: Partial<Style> = {};

  // --- Parsers ---

  const parseExactMatches = (cls: string): boolean => {
    if (fontSizes[cls]) {
      if (!style.font) style.font = {};
      style.font.size = fontSizes[cls];
      return true;
    }
    if (fontStyles[cls]) {
      if (!style.font) style.font = {};
      Object.assign(style.font, fontStyles[cls]);
      return true;
    }
    if (alignmentStyles[cls]) {
      if (!style.alignment) style.alignment = {};
      Object.assign(style.alignment, alignmentStyles[cls]);
      return true;
    }
    return false;
  };

  const parseTextColor = (cls: string): boolean => {
    if (cls.startsWith('text-')) {
      const colorKey = cls.substring(5);
      if (colors[colorKey]) {
        if (!style.font) style.font = {};
        style.font.color = { argb: colors[colorKey] };
        return true;
      }
    }
    return false;
  };

  const parseBackgroundColor = (cls: string): boolean => {
    if (cls.startsWith('bg-')) {
      const colorKey = cls.substring(3);
      if (colors[colorKey]) {
        style.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: colors[colorKey] },
        };
        return true;
      }
    }
    return false;
  };

  const parsers = [parseExactMatches, parseTextColor, parseBackgroundColor];

  // --- Border Info ---
  const borderInfo: {
    sides: BorderSide[];
    style?: Border['style'];
    color?: { argb: string };
  } = { sides: [] };

  // --- Processing Loop ---
  for (const cls of classes) {
    let recognized = false;
    // Handle borders separately since they are composed of multiple classes
    if (cls.startsWith('border')) {
      const parts = cls.split('-');
      if (borderSideMap[cls]) {
        borderInfo.sides.push(...borderSideMap[cls]);
        recognized = true;
      } else {
        const colorKey = parts.slice(1).join('-');
        if (colors[colorKey]) {
          borderInfo.color = { argb: colors[colorKey] };
          recognized = true;
        }
        if (borderStyles[parts[1]]) {
          borderInfo.style = borderStyles[parts[1]];
          recognized = true;
        }
      }
    }
    // Apply other parsers
    for (const parser of parsers) {
      if (parser(cls)) {
        recognized = true;
        break;
      }
    }
    // If not recognized by any parser, throw an error
    if (!recognized) {
      throw new Error(`[tailwindExcel] Unknown or unsupported class: '${cls}'`);
    }
  }

  // --- Finalize Styles ---
  if (borderInfo.sides.length > 0) {
    style.border = {};
    const borderStyle: Partial<Border> = {};
    borderStyle.style = borderInfo.style || 'thin';
    if (borderInfo.color) borderStyle.color = borderInfo.color;

    const uniqueSides = [...new Set(borderInfo.sides)];
    for (const side of uniqueSides) {
      (style.border as any)[side] = { ...borderStyle };
    }
  }

  return style;
}
