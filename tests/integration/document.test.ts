/**
 * Integration tests for document generation
 */

import { DocxGenerator } from '../../src/DocxGenerator';
import * as path from 'path';
import * as fs from 'fs';
import AdmZip from 'adm-zip';

const fixturesPath = path.join(__dirname, '..', 'fixtures');
const templatePath = path.join(fixturesPath, 'template.docx');
const outputPath = path.join(fixturesPath, 'output');

describe('Document Generation Integration', () => {
    beforeAll(() => {
        // Ensure output directory exists
        if (!fs.existsSync(outputPath)) {
            fs.mkdirSync(outputPath, { recursive: true });
        }
    });

    describe('Full document generation', () => {
        it('should generate valid DOCX from template', async () => {
            // Skip if fixture doesn't exist
            if (!fs.existsSync(templatePath)) {
                console.log('Skipping integration test: template fixture not found');
                return;
            }

            const generator = new DocxGenerator();
            await generator.loadTemplate(templatePath);

            generator.setData({
                name: 'Integration Test',
                date: '2026-01-09',
                company: 'Test Corp'
            });

            const buffer = await generator.generate();

            // Verify buffer is not empty
            expect(buffer.length).toBeGreaterThan(0);

            // Verify it's a valid ZIP (DOCX is a ZIP file)
            const zip = new AdmZip(buffer);
            const entries = zip.getEntries();
            expect(entries.length).toBeGreaterThan(0);

            // Verify required DOCX structure
            const entryNames = entries.map(e => e.entryName);
            expect(entryNames).toContain('[Content_Types].xml');
            expect(entryNames.some(name => name.includes('word/document.xml'))).toBe(true);
        });

        it('should save document to file', async () => {
            if (!fs.existsSync(templatePath)) {
                console.log('Skipping: template fixture not found');
                return;
            }

            const generator = new DocxGenerator();
            await generator.loadTemplate(templatePath);

            generator.setData({ name: 'Save Test' });

            const outputFile = path.join(outputPath, 'test-output.docx');
            await generator.save(outputFile);

            expect(fs.existsSync(outputFile)).toBe(true);

            // Cleanup
            fs.unlinkSync(outputFile);
        });

        it('should generate document with setData without errors', async () => {
            if (!fs.existsSync(templatePath)) {
                console.log('Skipping: template fixture not found');
                return;
            }

            const generator = new DocxGenerator();
            await generator.loadTemplate(templatePath);

            const testValue = `UniqueTestValue_${Date.now()}`;
            generator.setData({ name: testValue });

            const buffer = await generator.generate();

            // Verify document is generated
            expect(buffer.length).toBeGreaterThan(0);

            // Extract and verify document.xml exists
            const zip = new AdmZip(buffer);
            const documentEntry = zip.getEntry('word/document.xml');
            expect(documentEntry).not.toBeNull();
        });
    });

    describe('Table generation', () => {
        it('should generate document with addTable without errors', async () => {
            if (!fs.existsSync(templatePath)) {
                console.log('Skipping: template fixture not found');
                return;
            }

            const generator = new DocxGenerator();
            await generator.loadTemplate(templatePath);

            // Add table - even if placeholder doesn't exist, it shouldn't error
            generator.addTable({
                placeholder: 'itemsTable',
                headers: [
                    { name: 'Item', key: 'item', width: 2000 },
                    { name: 'Price', key: 'price', width: 2000 }
                ],
                rows: [
                    ['Widget A', '$99'],
                    ['Widget B', '$149']
                ]
            });

            const buffer = await generator.generate();

            // Verify document is generated successfully
            expect(buffer.length).toBeGreaterThan(0);

            const zip = new AdmZip(buffer);
            const documentEntry = zip.getEntry('word/document.xml');
            expect(documentEntry).not.toBeNull();
        });
    });
});
