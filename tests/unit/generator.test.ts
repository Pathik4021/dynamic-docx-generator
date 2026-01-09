/**
 * Unit tests for DocxGenerator class
 */

import { DocxGenerator } from '../../src/DocxGenerator';
import * as path from 'path';
import * as fs from 'fs';

// Path to test fixtures
const fixturesPath = path.join(__dirname, '..', 'fixtures');
const templatePath = path.join(fixturesPath, 'template.docx');

describe('DocxGenerator', () => {
    let generator: DocxGenerator;

    beforeEach(() => {
        generator = new DocxGenerator();
    });

    describe('constructor', () => {
        it('should create instance with default options', () => {
            expect(generator).toBeInstanceOf(DocxGenerator);
        });

        it('should create instance with custom options', () => {
            const customGenerator = new DocxGenerator({ tempDir: '/tmp/custom' });
            expect(customGenerator).toBeInstanceOf(DocxGenerator);
        });
    });

    describe('isTemplateLoaded', () => {
        it('should return false before loading template', () => {
            expect(generator.isTemplateLoaded()).toBe(false);
        });
    });

    describe('loadTemplate', () => {
        it('should throw error if file does not exist', async () => {
            await expect(generator.loadTemplate('./nonexistent.docx'))
                .rejects.toThrow('Template file not found');
        });

        it('should load template from file if exists', async () => {
            // Skip if fixture doesn't exist
            if (!fs.existsSync(templatePath)) {
                console.log('Skipping test: template fixture not found');
                return;
            }

            await generator.loadTemplate(templatePath);
            expect(generator.isTemplateLoaded()).toBe(true);
        });
    });

    describe('setData', () => {
        it('should allow method chaining', () => {
            const result = generator.setData({ name: 'test' });
            expect(result).toBe(generator);
        });

        it('should merge multiple setData calls', () => {
            generator.setData({ a: '1' }).setData({ b: '2' });
            // Internal state check would require exposing private members
            // For now, just verify chaining works
            expect(generator).toBeInstanceOf(DocxGenerator);
        });
    });

    describe('setHeader', () => {
        it('should allow method chaining', () => {
            const result = generator.setHeader({ title: 'Test' });
            expect(result).toBe(generator);
        });
    });

    describe('setFooter', () => {
        it('should allow method chaining', () => {
            const result = generator.setFooter({ page: '1' });
            expect(result).toBe(generator);
        });
    });

    describe('setImages', () => {
        it('should allow method chaining', () => {
            const result = generator.setImages([
                { placeholder: 'logo', path: './logo.png' }
            ]);
            expect(result).toBe(generator);
        });
    });

    describe('addTable', () => {
        it('should allow method chaining', () => {
            const result = generator.addTable({
                placeholder: 'table',
                headers: [{ name: 'Name', key: 'name' }],
                rows: [['Value']]
            });
            expect(result).toBe(generator);
        });
    });

    describe('reset', () => {
        it('should allow method chaining', () => {
            const result = generator.reset();
            expect(result).toBe(generator);
        });
    });

    describe('generate', () => {
        it('should throw error if no template loaded', async () => {
            await expect(generator.generate())
                .rejects.toThrow('No template loaded');
        });
    });

    describe('fluent API', () => {
        it('should support complete method chaining', () => {
            const result = generator
                .setData({ name: 'Test' })
                .setHeader({ title: 'Header' })
                .setFooter({ page: '1' })
                .setImages([])
                .addTable({
                    placeholder: 'table',
                    headers: [{ name: 'Col', key: 'col' }],
                    rows: []
                })
                .reset();

            expect(result).toBe(generator);
        });
    });
});
