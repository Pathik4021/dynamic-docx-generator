const { DocxGenerator } = require('../dist'); // Using local build for testing
const fs = require('fs');
const path = require('path');

async function generateDocument() {
  console.log('🚀 Starting comprehensive user scenario test...');

  const generator = new DocxGenerator();
  
  // 1. Load the template (created by run-test.js, or we assume it exists)
  // If it doesn't exist, we should probably warn or try to run run-test.js logic first.
  // For now, assuming test-template.docx exists from previous steps.
  const templatePath = './test-template.docx';
  if (!fs.existsSync(templatePath)) {
    console.error(`❌ Template not found at ${templatePath}. Please run 'node run-test.js' first to generate the template.`);
    return;
  }
  
  console.log('📂 Loading template...');
  await generator.loadTemplate(templatePath);
  
  // 2. Set Body Data
  console.log('📝 Setting body placeholders...');
  generator.setData({
    documentTitle: 'Comprehensive Status Report',
    companyName: 'Global Tech Solutions',
    productName: 'Dynamic DOCX Generator',
    date: new Date().toLocaleDateString(),
    description: 'This document demonstrates the full capabilities of the dynamic-docx-generator package, including tables, headers, footers, and text replacement.',
    preparedBy: 'Senior Developer'
  });

  // 3. Set Header Data
  // (Assuming template has header placeholders, if not, this won't break anything but won't show)
  // The test-template.docx from run-test.js didn't explicitly create header/footer parts with placeholders 
  // other than what's in document.xml, but the API supports it.
  // We'll call it to show usage.
  console.log('📝 Setting header/footer data...');
  generator.setHeader({
    headerTitle: 'Confidential Internal Document',
    headerDate: '2026'
  });

  generator.setFooter({
    footerText: 'Page 1 of 1 - Generated automatically'
  });

  // 4. Add Dynamic Table
  console.log('📊 Adding dynamic table...');
  generator.addTable({
    placeholder: 'itemsTable',
    headers: [
      { name: 'ID', key: 'id', width: 1000 },
      { name: 'Feature', key: 'feature', width: 3000 },
      { name: 'Status', key: 'status', width: 2000 },
      { name: 'Priority', key: 'priority', width: 1500 },
      { name: 'Owner', key: 'owner', width: 2000 }
    ],
    rows: [
      ['001', 'Placeholder Replacement', 'Completed', 'High', 'Team A'],
      ['002', 'Table Generation', 'Completed', 'Critical', 'Team B'],
      ['003', 'Image Insertion', 'In Progress', 'Medium', 'Team A'],
      ['004', 'Header/Footer Support', 'Completed', 'Low', 'Team C'],
      ['005', 'NPM Publishing', 'Pending', 'High', 'Team B']
    ],
    style: {
      headerBgColor: '4472C4', // Azure blue
      headerTextColor: 'FFFFFF', // White
      fontSize: 22, // 11pt
      fontFamily: 'Calibri',
      // New custom styling features
      headerHeight: 600,
      rowHeight: 400,
      tableAlign: 'center',
      cellPadding: 150,
      borderSize: 8
    }
  });

  // 5. Add Images (Optional - skipping if no image file present, or we could generate valid buffer)
  // To make this robust, we can use a small 1x1 pixel base64 PNG buffer
  const pixelBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
  const imageBuffer = Buffer.from(pixelBase64, 'base64');
  
  // Only add if there was an image placeholder in the template (our test-template.docx currently doesn't have one)
  // But we can show the API usage.
  /*
  generator.setImages([
    {
      placeholder: 'limitless_logo',
      source: imageBuffer,
      width: 100,
      height: 100
    }
  ]);
  */

  // 6. Save the generated document
  const outputPath = './output-user-scenario.docx';
  console.log(`💾 Saving document to ${outputPath}...`);
  await generator.save(outputPath);
  
  console.log('✅ Document generated successfully!');
  console.log(`📄 Path: ${path.resolve(outputPath)}`);
}

generateDocument().catch(console.error);
