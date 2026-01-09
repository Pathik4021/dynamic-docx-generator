/**
 * End-to-End Test for dynamic-docx-generator
 * 
 * This test creates a DOCX template from scratch with placeholders,
 * then uses the package to generate a document with dynamic content.
 * 
 * Run: node run-test.js
 */

// Use local build for testing
const { DocxGenerator } = require('../dist');
const AdmZip = require('adm-zip');
const fs = require('fs');
const path = require('path');

// Create a simple DOCX template with placeholders
function createTestTemplate() {
  console.log('📝 Creating test template with placeholders...');
  
  // A DOCX file is a ZIP archive with XML files inside
  const zip = new AdmZip();
  
  // [Content_Types].xml
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  zip.addFile('[Content_Types].xml', Buffer.from(contentTypes));
  
  // _rels/.rels
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  zip.addFile('_rels/.rels', Buffer.from(rels));
  
  // word/_rels/document.xml.rels
  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
  zip.addFile('word/_rels/document.xml.rels', Buffer.from(docRels));
  
  // word/document.xml - The main document with placeholders
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <!-- Title -->
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="48"/></w:rPr>
        <w:t>{{documentTitle}}</w:t>
      </w:r>
    </w:p>
    
    <!-- Company Info -->
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Company: </w:t></w:r>
      <w:r><w:t>{{companyName}}</w:t></w:r>
    </w:p>
    
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Product: </w:t></w:r>
      <w:r><w:t>{{productName}}</w:t></w:r>
    </w:p>
    
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Date: </w:t></w:r>
      <w:r><w:t>{{date}}</w:t></w:r>
    </w:p>
    
    <!-- Description -->
    <w:p>
      <w:r><w:t>{{description}}</w:t></w:r>
    </w:p>
    
    <!-- Empty paragraph before table -->
    <w:p/>
    
    <!-- Table Placeholder -->
    <w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Product List:</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>{{itemsTable}}</w:t></w:r>
    </w:p>
    
    <!-- Empty paragraph after table -->
    <w:p/>
    
    <!-- Footer section -->
    <w:p>
      <w:r><w:t>Document prepared by: {{preparedBy}}</w:t></w:r>
    </w:p>
    
    <w:sectPr/>
  </w:body>
</w:document>`;
  zip.addFile('word/document.xml', Buffer.from(document));
  
  // Save the template
  const templatePath = path.join(__dirname, 'test-template.docx');
  zip.writeZip(templatePath);
  console.log(`✅ Template created: ${templatePath}`);
  
  return templatePath;
}

// Run the end-to-end test
async function runE2ETest() {
  console.log('\n🧪 Running End-to-End Test for dynamic-docx-generator\n');
  console.log('='.repeat(50));
  
  try {
    // Step 1: Create test template
    const templatePath = createTestTemplate();
    
    // Step 2: Initialize generator
    console.log('\n📦 Initializing DocxGenerator...');
    const generator = new DocxGenerator();
    
    // Step 3: Load template
    console.log('📂 Loading template...');
    await generator.loadTemplate(templatePath);
    console.log('✅ Template loaded');
    
    // Step 4: Set placeholder data
    console.log('📝 Setting placeholder data...');
    generator.setData({
      documentTitle: 'Product Catalog 2026',
      companyName: 'TechCorp Industries',
      productName: 'Enterprise Suite',
      date: new Date().toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
      }),
      description: 'This document contains our complete product listing with specifications and pricing information.',
      preparedBy: 'Pathik - Sales Department'
    });
    console.log('✅ Data set');
    
    // Step 5: Add dynamic table
    console.log('📊 Adding dynamic table...');
    generator.addTable({
      placeholder: 'itemsTable',
      headers: [
        { name: 'Sr. No.', key: 'sr', width: 800 },
        { name: 'Product Name', key: 'name', width: 3000 },
        { name: 'Category', key: 'category', width: 2000 },
        { name: 'Price', key: 'price', width: 1500 },
        { name: 'Status', key: 'status', width: 1500 }
      ],
      rows: [
        ['1', 'Widget Pro', 'Hardware', '$299.00', 'In Stock'],
        ['2', 'Data Analyzer', 'Software', '$499.00', 'Available'],
        ['3', 'Cloud Storage 100GB', 'Service', '$9.99/mo', 'Active'],
        ['4', 'Premium Support', 'Service', '$199.00/yr', 'Active'],
        ['5', 'API Access', 'Software', '$49.00/mo', 'Available']
      ],
      style: {
        headerBgColor: '2E86AB',
        headerTextColor: 'FFFFFF',
        fontSize: 22,
        fontFamily: 'Arial'
      }
    });
    console.log('✅ Table added with 5 rows');
    
    // Step 6: Generate document
    console.log('⚙️  Generating document...');
    const buffer = await generator.generate();
    console.log(`✅ Document generated (${buffer.length} bytes)`);
    
    // Step 7: Save to file
    const outputPath = path.join(__dirname, 'output-e2e-test.docx');
    await generator.save(outputPath);
    console.log(`✅ Document saved: ${outputPath}`);
    
    // Step 8: Verify the output
    console.log('\n🔍 Verifying output...');
    const outputZip = new AdmZip(outputPath);
    const documentXml = outputZip.getEntry('word/document.xml');
    
    if (documentXml) {
      const content = documentXml.getData().toString('utf8');
      
      const checks = [
        { name: 'Company Name', found: content.includes('TechCorp Industries') },
        { name: 'Product Name', found: content.includes('Enterprise Suite') },
        { name: 'Document Title', found: content.includes('Product Catalog 2026') },
        { name: 'Table Structure', found: content.includes('<w:tbl>') },
        { name: 'Table Data (Widget Pro)', found: content.includes('Widget Pro') },
        { name: 'Table Data (Data Analyzer)', found: content.includes('Data Analyzer') },
        { name: 'Prepared By', found: content.includes('Pathik') }
      ];
      
      console.log('\nVerification Results:');
      let allPassed = true;
      for (const check of checks) {
        const status = check.found ? '✅' : '❌';
        console.log(`  ${status} ${check.name}`);
        if (!check.found) allPassed = false;
      }
      
      console.log('\n' + '='.repeat(50));
      if (allPassed) {
        console.log('🎉 ALL TESTS PASSED!');
        console.log(`\n📄 Open the document to verify:\n   ${outputPath}`);
      } else {
        console.log('⚠️  Some checks failed. Please review.');
      }
    }
    
  } catch (error) {
    console.error('\n❌ Test failed:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

// Run the test
runE2ETest();
