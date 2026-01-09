/**
 * XML manipulation utilities for DOCX processing
 */

import { replaceAll, fixDoubleEscaping } from './string';

/**
 * Find and replace content in XML string
 */
export const replaceXmlContent = (
    xml: string,
    search: string,
    replacement: string
): string => {
    let result = replaceAll(xml, search, replacement);
    result = fixDoubleEscaping(result);
    return result;
};

/**
 * Insert content before a specific XML tag
 */
export const insertBeforeTag = (
    xml: string,
    tag: string,
    content: string
): string => {
    const index = xml.indexOf(tag);
    if (index === -1) return xml;

    return xml.slice(0, index) + content + xml.slice(index);
};

/**
 * Insert content after a specific XML tag
 */
export const insertAfterTag = (
    xml: string,
    tag: string,
    content: string
): string => {
    const index = xml.indexOf(tag);
    if (index === -1) return xml;

    const endIndex = index + tag.length;
    return xml.slice(0, endIndex) + content + xml.slice(endIndex);
};

/**
 * Replace content between two tags
 */
export const replaceBetweenTags = (
    xml: string,
    startTag: string,
    endTag: string,
    replacement: string
): string => {
    const startIndex = xml.indexOf(startTag);
    const endIndex = xml.indexOf(endTag, startIndex);

    if (startIndex === -1 || endIndex === -1) return xml;

    return (
        xml.slice(0, startIndex + startTag.length) +
        replacement +
        xml.slice(endIndex)
    );
};

/**
 * Add relationship to document.xml.rels
 */
export const addRelationship = (
    relsContent: string,
    relationshipXml: string
): string => {
    const closingTag = '</Relationships>';
    return replaceAll(relsContent, closingTag, relationshipXml + '\n' + closingTag);
};

/**
 * Extract relationship IDs from rels XML
 */
export const extractRelationshipIds = (relsContent: string): string[] => {
    const matches = relsContent.match(/Id="([^"]+)"/g);
    if (!matches) return [];

    return matches.map(m => m.replace('Id="', '').replace('"', ''));
};

/**
 * Generate unique relationship ID not present in existing rels
 */
export const generateUniqueRelId = (existingIds: string[]): string => {
    let id = 1;
    while (existingIds.includes(`rId${id}`)) {
        id++;
    }
    return `rId${id}`;
};

/**
 * Parse content types from [Content_Types].xml
 */
export const addContentType = (
    contentTypesXml: string,
    extension: string,
    contentType: string
): string => {
    const existing = `Extension="${extension}"`;
    if (contentTypesXml.includes(existing)) {
        return contentTypesXml;
    }

    const newType = `<Default Extension="${extension}" ContentType="${contentType}" />`;
    const closingTag = '</Types>';
    return replaceAll(contentTypesXml, closingTag, newType + '\n' + closingTag);
};

/**
 * Common content types for images
 */
export const IMAGE_CONTENT_TYPES: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp'
};

/**
 * Get content type for image extension
 */
export const getImageContentType = (extension: string): string => {
    const ext = extension.toLowerCase().replace('.', '');
    return IMAGE_CONTENT_TYPES[ext] || 'image/png';
};

/**
 * Validate XML structure (basic check)
 */
export const isValidXml = (xml: string): boolean => {
    try {
        // Basic check: matching opening/closing tags
        const openTags = (xml.match(/<[^/][^>]*[^/]>/g) || []).length;
        const closeTags = (xml.match(/<\/[^>]+>/g) || []).length;
        const selfClose = (xml.match(/<[^>]+\/>/g) || []).length;

        // This is a rough approximation
        return Math.abs(openTags - closeTags - selfClose) < 5;
    } catch {
        return false;
    }
};
