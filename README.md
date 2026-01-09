# dynamic-docx-generator

Generate dynamic DOCX documents from templates with placeholder replacement, tables, and images.

[![npm version](https://badge.fury.io/js/dynamic-docx-generator.svg)](https://www.npmjs.com/package/dynamic-docx-generator)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- 📝 **Template-based generation** - Use existing DOCX files as templates
- 🔄 **Placeholder replacement** - Replace `{{placeholder}}` tags with dynamic data
- 📊 **Dynamic tables** - Generate tables from data arrays
- 🖼️ **Image insertion** - Embed images from files, URLs, or buffers
- 📄 **Header/Footer support** - Customize document headers and footers
- ⚡ **Fluent API** - Chain methods for clean, readable code
- 📦 **Zero system dependencies** - No external tools like LibreOffice needed
- 🔧 **TypeScript support** - Full type definitions included

## Installation

```bash
npm install dynamic-docx-generator
```

## Quick Start

```javascript
const { DocxGenerator } = require('dynamic-docx-generator');

async function generateDocument() {
  const generator = new DocxGenerator();
  
  // Load a DOCX template
  await generator.loadTemplate('./template.docx');
  
  // Replace placeholders
  generator.setData({
    companyName: 'Acme Corporation',
    productName: 'Widget Pro',
    date: '2026-01-09'
  });
  
  // Save the generated document
  await generator.save('./output.docx');
  
  console.log('Document generated successfully!');
}

generateDocument();
```

## API Reference

### DocxGenerator

The main class for document generation.

#### Constructor

```typescript
const generator = new DocxGenerator(options?: GeneratorOptions);
```

**Options:**
- `tempDir?: string` - Custom temporary directory for processing

#### Methods

##### loadTemplate(source)

Load a DOCX template from a file path or buffer.

```typescript
await generator.loadTemplate('./template.docx');
// or
await generator.loadTemplate(buffer);
```

##### setData(data)

Set placeholder values for the document body.

```typescript
generator.setData({
  name: 'John Doe',
  company: 'Acme Corp',
  date: new Date().toLocaleDateString()
});
```

##### setHeader(data)

Set placeholder values for document headers.

```typescript
generator.setHeader({
  documentTitle: 'Annual Report',
  version: '1.0'
});
```

##### setFooter(data)

Set placeholder values for document footers.

```typescript
generator.setFooter({
  copyright: '© 2026 Acme Corp'
});
```

##### setImages(images)

Add images to be inserted into the document.

```typescript
generator.setImages([
  {
    placeholder: 'logo',
    path: './logo.png',
    width: 914400  // In EMUs (914400 EMUs = 1 inch)
  },
  {
    placeholder: 'signature',
    url: 'https://example.com/signature.png',
    width: 457200,
    height: 228600
  }
]);
```

##### addTable(config)

Add a dynamic table to the document.

```typescript
generator.addTable({
  placeholder: 'itemsTable',
  headers: [
    { name: 'Item', key: 'item', width: 3000 },
    { name: 'Quantity', key: 'qty', width: 1500 },
    { name: 'Price', key: 'price', width: 1500 }
  ],
  rows: [
    ['Widget A', '10', '$99.00'],
    ['Widget B', '5', '$149.00'],
    ['Widget C', '3', '$199.00']
  ],
  style: {
    // Basic Colors
    headerBgColor: 'C9DAF8',
    headerTextColor: '000000',
    borderColor: 'auto',
    
    // Fonts
    fontSize: 24,       // 12pt
    fontFamily: 'Calibri',
    
    // Layout & Dimensions (New!)
    headerHeight: 400,  // Header row height in twips
    rowHeight: 350,     // Data row height in twips
    tableAlign: 'center', // 'center', 'left', or 'right'
    cellPadding: 100,   // Cell margins in twips
    borderSize: 4       // Border thickness (eights of a point)
  }
});
```

### Table Style Properties

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `headerBgColor` | string | `'C9DAF8'` | Hex color for header background (no #) |
| `headerTextColor` | string | `'000000'` | Hex color for header text (no #) |
| `borderColor` | string | `'auto'` | Hex color for borders (no #) |
| `fontSize` | number | `24` | Font size in half-points (24 = 12pt) |
| `fontFamily` | string | `'Times New Roman'` | Font family name |
| `headerHeight` | number | `400` | Height of header row in twips |
| `rowHeight` | number | `350` | Height of data rows in twips |
| `tableAlign` | string | `'center'` | Alignment of the table: `'center'`, `'left'`, `'right'` |
| `cellPadding` | number | `100` | Padding inside cells in twips |
| `borderSize` | number | `4` | Thickness of borders in eighths of a point |


##### generate()

Generate the document and return as a Buffer.

```typescript
const buffer = await generator.generate();
```

##### save(outputPath)

Generate and save the document to a file.

```typescript
await generator.save('./output.docx');
```

##### reset()

Reset generator state while keeping the template loaded.

```typescript
generator.reset();
```

## Template Format

Create your template in Microsoft Word or any DOCX-compatible editor:

1. Add placeholders using `{{placeholderName}}` syntax
2. For tables, add a placeholder like `{{itemsTable}}`
3. For images, add a placeholder like `{{logo}}`

**Example template content:**

```
Document Title: {{documentTitle}}
Company: {{companyName}}
Date: {{date}}

{{itemsTable}}

Signature: {{signature}}
```

## Image Units (EMUs)

DOCX uses English Metric Units (EMUs) for dimensions:
- **914400 EMUs = 1 inch**
- **360000 EMUs = 1 cm**
- **9525 EMUs = 1 pixel (at 96 DPI)**

Helper functions are available:

```typescript
const { ImageUtils } = require('dynamic-docx-generator');

// Convert to EMUs
const widthInEmu = ImageUtils.pixelsToEmu(200);  // 200px -> EMUs
const heightInEmu = ImageUtils.inchesToEmu(2);   // 2 inches -> EMUs
const sizeInEmu = ImageUtils.cmToEmu(5);         // 5cm -> EMUs
```

## Complete Example

```javascript
const { DocxGenerator, ImageUtils } = require('dynamic-docx-generator');
const path = require('path');

async function createInvoice() {
  const generator = new DocxGenerator();
  
  await generator.loadTemplate('./invoice-template.docx');
  
  // Company details
  generator.setData({
    invoiceNumber: 'INV-2026-001',
    date: '2026-01-09',
    customerName: 'John Doe',
    customerAddress: '123 Main Street, City',
    total: '$447.00',
    tax: '$44.70',
    grandTotal: '$491.70'
  });
  
  // Header
  generator.setHeader({
    companyName: 'Acme Corp'
  });
  
  // Company logo
  generator.setImages([
    {
      placeholder: 'companyLogo',
      path: './assets/logo.png',
      width: ImageUtils.inchesToEmu(1.5)
    }
  ]);
  
  // Line items table
  generator.addTable({
    placeholder: 'lineItems',
    headers: [
      { name: 'Description', key: 'desc', width: 4000 },
      { name: 'Qty', key: 'qty', width: 1000 },
      { name: 'Unit Price', key: 'price', width: 1500 },
      { name: 'Total', key: 'total', width: 1500 }
    ],
    rows: [
      ['Widget Pro', '2', '$99.00', '$198.00'],
      ['Widget Plus', '1', '$149.00', '$149.00'],
      ['Premium Support', '1', '$100.00', '$100.00']
    ]
  });
  
  await generator.save('./generated-invoice.docx');
  console.log('Invoice generated!');
}

createInvoice();
```

## TypeScript Usage

```typescript
import { DocxGenerator, ImageConfig, TableConfig } from 'dynamic-docx-generator';

async function generateReport(): Promise<void> {
  const generator = new DocxGenerator();
  
  await generator.loadTemplate('./template.docx');
  
  const images: ImageConfig[] = [
    { placeholder: 'chart', path: './chart.png', width: 914400 }
  ];
  
  const table: TableConfig = {
    placeholder: 'dataTable',
    headers: [
      { name: 'Name', key: 'name', width: 3000 },
      { name: 'Value', key: 'value', width: 2000 }
    ],
    rows: [['Item 1', '100'], ['Item 2', '200']]
  };
  
  generator
    .setData({ title: 'Monthly Report' })
    .setImages(images)
    .addTable(table);
  
  await generator.save('./report.docx');
}
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT © 2026
