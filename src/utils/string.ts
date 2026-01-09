/**
 * String manipulation utilities for DOCX template processing
 */

/**
 * Escape special characters for use in regular expressions
 */
export const escapeRegExp = (string: string): string => {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

/**
 * Replace all occurrences of a string (case-sensitive)
 */
export const replaceAll = (str: string, search: string, replacement: string): string => {
    if (typeof str !== 'string') {
        return str;
    }
    return str.replace(new RegExp(escapeRegExp(search), 'g'), replacement);
};

/**
 * Replace placeholder with value, handling XML-safe escaping
 * Supports both {{placeholder}} and <w:t>placeholder</w:t> formats
 */
export const replacePlaceholder = (
    content: string,
    placeholder: string,
    value: string
): string => {
    let result = content;

    // Escape the value for XML
    const safeValue = escapeForXml(value);

    // Replace {{placeholder}} format
    result = replaceAll(result, `{{${placeholder}}}`, safeValue);

    // Replace <w:t>placeholder</w:t> format (Word XML)
    result = replaceAll(result, `<w:t>${placeholder}</w:t>`, `<w:t>${safeValue}</w:t>`);

    // Replace when placeholder is split across XML tags (common in Word)
    // Pattern: <w:t>place</w:t></w:r><w:r><w:t>holder</w:t>
    const placeholderPattern = placeholder.split('').join('[^a-zA-Z]*');
    const regex = new RegExp(`<w:t[^>]*>\\s*${placeholderPattern}\\s*</w:t>`, 'gi');
    result = result.replace(regex, `<w:t>${safeValue}</w:t>`);

    return result;
};

/**
 * Escape special characters for XML content
 */
export const escapeForXml = (text: string): string => {
    if (!text || typeof text !== 'string') return '';

    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
};

/**
 * Unescape XML entities back to regular characters
 */
export const unescapeXml = (text: string): string => {
    if (!text || typeof text !== 'string') return '';

    return text
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
};

/**
 * Fix double-escaped XML entities
 */
export const fixDoubleEscaping = (content: string): string => {
    let result = content;
    result = replaceAll(result, '&amp;amp;', '&amp;');
    result = replaceAll(result, '&amp;lt;', '&lt;');
    result = replaceAll(result, '&amp;gt;', '&gt;');
    result = replaceAll(result, '&amp;quot;', '&quot;');
    result = replaceAll(result, '&amp;apos;', '&apos;');
    return result;
};

/**
 * Strip HTML tags and return plain text
 */
export const stripHtml = (html: string): string => {
    if (!html || typeof html !== 'string') return '';

    return html
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
};

/**
 * Generate a unique ID for image relationships
 */
export const generateImageId = (): string => {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const nums = '0123456789';

    let id = 'rId';
    for (let i = 0; i < 2; i++) {
        id += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    for (let i = 0; i < 4; i++) {
        id += nums.charAt(Math.floor(Math.random() * nums.length));
    }

    return id;
};

/**
 * Check if a string contains any of the given placeholders
 */
export const containsPlaceholder = (content: string, placeholders: string[]): boolean => {
    return placeholders.some(p =>
        content.includes(`{{${p}}}`) || content.includes(`<w:t>${p}</w:t>`)
    );
};

/**
 * Extract all placeholders from content (both {{}} and raw formats)
 */
export const extractPlaceholders = (content: string): string[] => {
    const placeholders: Set<string> = new Set();

    // Match {{placeholder}} format
    const bracketMatches = content.match(/\{\{([^}]+)\}\}/g);
    if (bracketMatches) {
        bracketMatches.forEach(match => {
            placeholders.add(match.slice(2, -2));
        });
    }

    return Array.from(placeholders);
};

/**
 * Wrap text with XML preserve space attribute for proper whitespace handling
 */
export const wrapWithPreserveSpace = (text: string): string => {
    return `<w:t xml:space="preserve">${escapeForXml(text)}</w:t>`;
};
