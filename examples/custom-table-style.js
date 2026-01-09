const { DocxGenerator } = require('../dist');
const path = require('path');
const fs = require('fs');

async function createStyledTableDocument() {
  console.log('🚀 Starting custom table style example...');

  const generator = new DocxGenerator();
  
  // 1. Load a simple template
  // If you don't have one, you can create a simple one with valid placeholders
  const templatePath = path.join(__dirname, '../tests/fixtures/simple-template.docx');
  
  if (!fs.existsSync(templatePath)) {
    console.error(`❌ Template not found at ${templatePath}. Please run specific tests to generate it or provide your own.`);
    // Fallback for standalone run if fixture doesn't exist in user env
    console.log('⚠️  Skipping execution as template is missing. Access this feature via your own template.');
    return;
  }

  await generator.loadTemplate(templatePath);
  
  // 2. Add a table with CUSTOM STYLES
  console.log('📊 Adding table with custom styles...');
  
  generator.addTable({
    placeholder: 'itemsTable',
    headers: [
        { name: 'Rank', key: 'rank', width: 1000 },
        { name: 'Language', key: 'lang', width: 3000 },
        { name: 'Rating', key: 'rating', width: 1500 }
    ],
    rows: [
        ['1', 'TypeScript', '⭐⭐⭐⭐⭐'],
        ['2', 'JavaScript', '⭐⭐⭐⭐'],
        ['3', 'Python', '⭐⭐⭐⭐⭐']
    ],
    // This is the new feature:
    style: {
      // Colors
      headerBgColor: '4472C4',   // Azure blue
      headerTextColor: 'FFFFFF', // White
      borderColor: 'E7E6E6',     // Light gray
      
      // Fonts
      fontFamily: 'Segoe UI',
      fontSize: 22,              // 11pt
      
      // Dimensions
      headerHeight: 600,         // Taller header
      rowHeight: 400,            // Taller rows
      
      // Layout
      tableAlign: 'center',      // Centered table
      cellPadding: 200,          // More padding
      borderSize: 8              // Thicker borders
    }
  });
  
  // 3. Save
  const outputPath = path.join(__dirname, 'output-styled-table.docx');
  await generator.save(outputPath);
  
  console.log(`✅ Document saved to: ${outputPath}`);
}

createStyledTableDocument().catch(console.error);
