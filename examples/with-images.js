/**
 * Example with images for dynamic-docx-generator
 * 
 * Run this example:
 * 1. npm run build
 * 2. node examples/with-images.js
 */

const { DocxGenerator, ImageUtils } = require('../dist');
const path = require('path');

async function main() {
  try {
    const generator = new DocxGenerator();

    const templatePath = path.join(__dirname, '..', 'tests', 'fixtures', 'template.docx');
    await generator.loadTemplate(templatePath);

    // Set basic data
    generator.setData({
      companyName: 'Tech Solutions Inc.',
      productName: 'Cloud Platform',
      date: new Date().toLocaleDateString()
    });

    // Add images (if you have image files)
    // Uncomment and adjust paths when you have actual images:
    /*
    generator.setImages([
      {
        placeholder: 'logo',
        path: './assets/logo.png',
        width: ImageUtils.inchesToEmu(1.5),
        height: ImageUtils.inchesToEmu(0.5)
      },
      {
        placeholder: 'signature',
        url: 'https://example.com/signature.png',
        width: ImageUtils.pixelsToEmu(200)
      }
    ]);
    */

    // You can also use buffers:
    /*
    const fs = require('fs');
    const logoBuffer = fs.readFileSync('./logo.png');
    generator.setImages([
      {
        placeholder: 'logo',
        buffer: logoBuffer,
        width: ImageUtils.cmToEmu(4)
      }
    ]);
    */

    const outputPath = path.join(__dirname, 'output-with-images.docx');
    await generator.save(outputPath);

    console.log('✅ Document with images generated!');
    console.log(`📄 Output: ${outputPath}`);

  } catch (error) {
    console.error('❌ Error:', error.message);
    process.exit(1);
  }
}

main();
