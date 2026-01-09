/**
 * DocxGenerator - Main class for dynamic DOCX document generation
 * 
 * @example
 * ```typescript
 * const generator = new DocxGenerator();
 * await generator.loadTemplate('./template.docx');
 * generator.setData({ name: 'John', company: 'Acme' });
 * await generator.save('./output.docx');
 * ```
 */

import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import AdmZip from 'adm-zip';

import {
    GeneratorOptions,
    PlaceholderData,
    ImageConfig,
    TableConfig,
    TableStyle
} from './types';

import { replacePlaceholder, replaceAll, fixDoubleEscaping, escapeRegExp } from './utils/string';
import { replaceXmlContent, addRelationship, extractRelationshipIds, addContentType, getImageContentType } from './utils/xml';
import { prepareImage, PreparedImage } from './utils/image';
import { generateTable, DEFAULT_TABLE_STYLE } from './utils/constants';

export class DocxGenerator {
    private options: GeneratorOptions;
    private zip: AdmZip | null = null;
    private templateLoaded: boolean = false;
    private data: PlaceholderData = {};
    private headerData: PlaceholderData = {};
    private footerData: PlaceholderData = {};
    private images: ImageConfig[] = [];
    private tables: TableConfig[] = [];
    private tempDir: string;

    /**
     * Create a new DocxGenerator instance
     * @param options - Configuration options
     */
    constructor(options: GeneratorOptions = {}) {
        this.options = options;
        this.tempDir = options.tempDir || os.tmpdir();
    }

    /**
     * Load a DOCX template from file path or buffer
     * @param source - Path to template file or Buffer containing template
     * @returns this for method chaining
     */
    async loadTemplate(source: string | Buffer): Promise<this> {
        try {
            if (typeof source === 'string') {
                // Load from file path
                if (!fs.existsSync(source)) {
                    throw new Error(`Template file not found: ${source}`);
                }
                this.zip = new AdmZip(source);
            } else {
                // Load from buffer
                this.zip = new AdmZip(source);
            }

            this.templateLoaded = true;
            return this;
        } catch (error) {
            throw new Error(`Failed to load template: ${(error as Error).message}`);
        }
    }

    /**
     * Set placeholder data for document body
     * @param data - Key-value pairs where key is placeholder name and value is replacement text
     * @returns this for method chaining
     * 
     * @example
     * ```typescript
     * generator.setData({
     *   companyName: 'Acme Corp',
     *   productName: 'Widget Pro',
     *   date: '2026-01-09'
     * });
     * ```
     */
    setData(data: PlaceholderData): this {
        this.data = { ...this.data, ...data };
        return this;
    }

    /**
     * Set placeholder data for document header
     * @param data - Key-value pairs for header placeholders
     * @returns this for method chaining
     */
    setHeader(data: PlaceholderData): this {
        this.headerData = { ...this.headerData, ...data };
        return this;
    }

    /**
     * Set placeholder data for document footer
     * @param data - Key-value pairs for footer placeholders
     * @returns this for method chaining
     */
    setFooter(data: PlaceholderData): this {
        this.footerData = { ...this.footerData, ...data };
        return this;
    }

    /**
     * Add images to be inserted into the document
     * @param images - Array of image configurations
     * @returns this for method chaining
     * 
     * @example
     * ```typescript
     * generator.setImages([
     *   { placeholder: '{{logo}}', path: './logo.png', width: 200 }
     * ]);
     * ```
     */
    setImages(images: ImageConfig[]): this {
        this.images = [...this.images, ...images];
        return this;
    }

    /**
     * Add a dynamic table to the document
     * @param config - Table configuration
     * @returns this for method chaining
     * 
     * @example
     * ```typescript
     * generator.addTable({
     *   placeholder: '{{itemsTable}}',
     *   headers: [
     *     { name: 'Item', key: 'item', width: 3000 },
     *     { name: 'Quantity', key: 'qty', width: 1500 }
     *   ],
     *   rows: [['Widget A', '10'], ['Widget B', '5']]
     * });
     * ```
     */
    addTable(config: TableConfig): this {
        this.tables.push(config);
        return this;
    }

    /**
     * Process placeholders in XML content
     */
    private processPlaceholders(content: string, data: PlaceholderData): string {
        let result = content;

        for (const [key, value] of Object.entries(data)) {
            result = replacePlaceholder(result, key, value);
        }

        // Fix any double escaping issues
        result = fixDoubleEscaping(result);

        return result;
    }

    /**
     * Process tables in XML content
     */
    private processTables(content: string): string {
        let result = content;

        for (const table of this.tables) {
            const headers = table.headers.map(h => ({
                name: h.name,
                width: h.width
            }));

            const style = {
                ...DEFAULT_TABLE_STYLE,
                ...table.style
            };

            const tableXml = generateTable(headers, table.rows, style);

            // Escape placeholder for regex
            const escapedPlaceholder = escapeRegExp(table.placeholder);

            // 1. Try to replace wrapping paragraph for {{placeholder}}
            // Matches: <w:p ...> ... {{placeholder}} ... </w:p>
            const placeholderRegex = new RegExp(`<w:p[^>]*>(?:(?!<\\/w:p>).)*?{{${escapedPlaceholder}}}(?:(?!<\\/w:p>).)*?<\\/w:p>`, 'gs');

            if (placeholderRegex.test(result)) {
                result = result.replace(placeholderRegex, tableXml);
            } else {
                // Fallback: simple replacement
                result = replaceAll(result, `{{${table.placeholder}}}`, tableXml);

                // Also try replacing raw placeholder if provided
                result = replaceAll(result, table.placeholder, tableXml);
            }
        }

        return result;
    }

    /**
     * Process images and add them to the document
     */
    private async processImages(): Promise<PreparedImage[]> {
        if (!this.zip || this.images.length === 0) {
            return [];
        }

        // Get existing relationship IDs
        const relsEntry = this.zip.getEntry('word/_rels/document.xml.rels');
        const relsContent = relsEntry ? relsEntry.getData().toString('utf8') : '';
        const existingIds = extractRelationshipIds(relsContent);

        // Prepare all images
        const preparedImages: PreparedImage[] = [];
        for (let i = 0; i < this.images.length; i++) {
            const prepared = await prepareImage(this.images[i], i, [
                ...existingIds,
                ...preparedImages.map(p => p.id)
            ]);
            preparedImages.push(prepared);
        }

        return preparedImages;
    }

    /**
     * Update document relationships with images
     */
    private updateRelationships(images: PreparedImage[]): void {
        if (!this.zip || images.length === 0) return;

        // Update document.xml.rels
        const relsPath = 'word/_rels/document.xml.rels';
        const relsEntry = this.zip.getEntry(relsPath);

        if (relsEntry) {
            let relsContent = relsEntry.getData().toString('utf8');

            for (const image of images) {
                relsContent = addRelationship(relsContent, image.relationshipXml);
            }

            this.zip.updateFile(relsPath, Buffer.from(relsContent, 'utf8'));
        }

        // Update [Content_Types].xml for image types
        const contentTypesPath = '[Content_Types].xml';
        const contentTypesEntry = this.zip.getEntry(contentTypesPath);

        if (contentTypesEntry) {
            let contentTypesXml = contentTypesEntry.getData().toString('utf8');

            const extensions = new Set(images.map(img => img.extension));
            for (const ext of extensions) {
                contentTypesXml = addContentType(
                    contentTypesXml,
                    ext,
                    getImageContentType(ext)
                );
            }

            this.zip.updateFile(contentTypesPath, Buffer.from(contentTypesXml, 'utf8'));
        }
    }

    /**
     * Add image files to the media folder
     */
    private addImageFiles(images: PreparedImage[]): void {
        if (!this.zip || images.length === 0) return;

        // Ensure media folder exists
        for (const image of images) {
            this.zip.addFile(`word/media/${image.fileName}`, image.buffer);
        }
    }

    /**
     * Replace image placeholders in document content
     */
    private replaceImagePlaceholders(content: string, images: PreparedImage[]): string {
        let result = content;

        for (const image of images) {
            // Replace {{placeholder}} with image XML
            result = replaceAll(result, image.placeholder, image.inlineXml);
            result = replaceAll(result, `{{${image.placeholder}}}`, image.inlineXml);
        }

        return result;
    }

    /**
     * Generate the final DOCX document as a Buffer
     * @returns Buffer containing the generated DOCX file
     */
    async generate(): Promise<Buffer> {
        if (!this.templateLoaded || !this.zip) {
            throw new Error('No template loaded. Call loadTemplate() first.');
        }

        try {
            // Process images first
            const preparedImages = await this.processImages();

            // Process document.xml
            const documentPath = 'word/document.xml';
            const documentEntry = this.zip.getEntry(documentPath);

            if (documentEntry) {
                let documentXml = documentEntry.getData().toString('utf8');

                // Replace placeholders
                documentXml = this.processPlaceholders(documentXml, this.data);

                // Replace tables
                documentXml = this.processTables(documentXml);

                // Replace image placeholders
                documentXml = this.replaceImagePlaceholders(documentXml, preparedImages);

                this.zip.updateFile(documentPath, Buffer.from(documentXml, 'utf8'));
            }

            // Process header files
            const headerFiles = this.zip.getEntries()
                .filter(entry => entry.entryName.match(/word\/header\d+\.xml/));

            for (const headerEntry of headerFiles) {
                let headerXml = headerEntry.getData().toString('utf8');
                headerXml = this.processPlaceholders(headerXml, { ...this.data, ...this.headerData });
                headerXml = this.replaceImagePlaceholders(headerXml, preparedImages);
                this.zip.updateFile(headerEntry.entryName, Buffer.from(headerXml, 'utf8'));
            }

            // Process footer files
            const footerFiles = this.zip.getEntries()
                .filter(entry => entry.entryName.match(/word\/footer\d+\.xml/));

            for (const footerEntry of footerFiles) {
                let footerXml = footerEntry.getData().toString('utf8');
                footerXml = this.processPlaceholders(footerXml, { ...this.data, ...this.footerData });
                footerXml = this.replaceImagePlaceholders(footerXml, preparedImages);
                this.zip.updateFile(footerEntry.entryName, Buffer.from(footerXml, 'utf8'));
            }

            // Add images to document
            if (preparedImages.length > 0) {
                this.updateRelationships(preparedImages);
                this.addImageFiles(preparedImages);
            }

            // Generate the final buffer
            return this.zip.toBuffer();
        } catch (error) {
            throw new Error(`Failed to generate document: ${(error as Error).message}`);
        }
    }

    /**
     * Save the generated document to a file
     * @param outputPath - Path where the document should be saved
     */
    async save(outputPath: string): Promise<void> {
        const buffer = await this.generate();

        // Ensure directory exists
        const dir = path.dirname(outputPath);
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
        }

        await fs.promises.writeFile(outputPath, buffer);
    }

    /**
     * Reset the generator state (keeps template loaded)
     */
    reset(): this {
        this.data = {};
        this.headerData = {};
        this.footerData = {};
        this.images = [];
        this.tables = [];
        return this;
    }

    /**
     * Check if a template has been loaded
     */
    isTemplateLoaded(): boolean {
        return this.templateLoaded;
    }
}
