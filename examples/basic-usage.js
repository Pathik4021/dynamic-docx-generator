/**
 * Basic usage example for dynamic-docx-generator
 * 
 * Run this example:
 * 1. npm run build
 * 2. node examples/basic-usage.js
 */

const { DocxGenerator } = require('../dist');
const path = require('path');

async function main() {
  try {
    // Create a new generator instance
    const generator = new DocxGenerator();

    // Load the template
    const templatePath = path.join(__dirname, '..', 'tests', 'fixtures', 'template.docx');
    await generator.loadTemplate(templatePath);

    // Set placeholder data
    generator.setData({
      companyName: 'Acme Corporation',
      productName: 'Widget Pro Max',
      date: new Date().toLocaleDateString(),
      documentNumber: 'DOC-2026-001',
      revisionNumber: '1.0'
    });

    // Set header data
    generator.setHeader({
      headerTitle: 'Product Documentation'
    });

    // Add a dynamic table
    generator.addTable({
      placeholder: 'itemsTable',
      headers: [
        { name: 'Feature', key: 'feature', width: 3000 },
        { name: 'Description', key: 'description', width: 4000 },
        { name: 'Status', key: 'status', width: 1500 }
      ],
      rows: [
        ['Auto-sync', 'Automatically syncs data across devices', 'Active'],
        ['Cloud Storage', '100GB cloud storage included', 'Active'],
        ['Premium Support', '24/7 premium support access', 'Active'],
        ['Analytics', 'Advanced analytics dashboard', 'Beta']
      ],
      style: {
        headerBgColor: '4A90D9',
        headerTextColor: 'FFFFFF',
        fontSize: 22
      }
    });

    // Save the generated document
    const outputPath = path.join(__dirname, 'output-basic.docx');
    await generator.save(outputPath);

    console.log('✅ Document generated successfully!');
    console.log(`📄 Output: ${outputPath}`);

  } catch (error) {
    console.error('❌ Error:', error.message);
    process.exit(1);
  }
}

main();
