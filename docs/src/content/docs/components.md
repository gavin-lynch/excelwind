---
title: Components
description: The JSX building blocks for Excelwind.
---

## Workbook
Root container for a file.
```tsx
<Workbook>...</Workbook>
```

## Worksheet
Defines a worksheet in the workbook.
```tsx
<Worksheet name="Sheet1" properties={{ tabColor: { argb: "FF0000" } }}>
  ...
</Worksheet>
```

## Column
Configures column width, format, or named ranges.
```tsx
<Column width={20} format='"$"#,##0.00' className="text-right" />
<Column id="Dates" width={15} format="yyyy-mm-dd" className="text-center" />
```

## Row
Defines a row of cells.
```tsx
<Row height={24} className="bg-gray-50">
  <Cell value="Hello" className="font-bold" />
</Row>
```

## Cell
Individual cell with value, format, spans, and optional images.
```tsx
<Cell value="Text" className="text-left" />
<Cell value={123} format='"$"#,##0.00' className="text-right" />
<Cell value="Merged" colSpan={2} rowSpan={2} className="text-center" />
```

## Group
Group rows or cells to share styles or processors.
```tsx
<Group className="bg-gray-100">...
</Group>
```

## Template
Loads an .xlsx file and expands data placeholders.
```tsx
<Template src="template.xlsx" data={{ columns: [], rows: [] }} />
```

## Image
Embeds an image either at worksheet level or inside a cell.
```tsx
<Image src="./logo.png" extension="png" range="A1:C3" />
```
