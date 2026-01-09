/**
 * dynamic-docx-generator
 * 
 * Generate dynamic DOCX documents from templates with placeholder replacement,
 * tables, and images.
 * 
 * @packageDocumentation
 */

// Main class
export { DocxGenerator } from './DocxGenerator';

// Types
export * from './types';

// Utilities (for advanced usage)
export * as StringUtils from './utils/string';
export * as XmlUtils from './utils/xml';
export * as ImageUtils from './utils/image';
export {
    generateTable,
    generateInlineImage,
    DEFAULT_TABLE_STYLE,
    escapeXml
} from './utils/constants';
