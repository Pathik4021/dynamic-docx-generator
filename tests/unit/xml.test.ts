/**
 * Unit tests for XML utilities
 */

import {
    replaceXmlContent,
    insertBeforeTag,
    insertAfterTag,
    addRelationship,
    extractRelationshipIds,
    generateUniqueRelId,
    addContentType,
    getImageContentType
} from '../../src/utils/xml';

describe('XML Utilities', () => {
    describe('replaceXmlContent', () => {
        it('should replace content in XML', () => {
            const xml = '<p>Hello World</p>';
            const result = replaceXmlContent(xml, 'World', 'Universe');
            expect(result).toBe('<p>Hello Universe</p>');
        });

        it('should fix double escaping', () => {
            const xml = '<p>A &amp;amp; B</p>';
            const result = replaceXmlContent(xml, 'nothing', 'nothing');
            expect(result).toBe('<p>A &amp; B</p>');
        });
    });

    describe('insertBeforeTag', () => {
        it('should insert content before tag', () => {
            const xml = '<root><child/></root>';
            const result = insertBeforeTag(xml, '<child/>', '<inserted/>');
            expect(result).toBe('<root><inserted/><child/></root>');
        });

        it('should return original if tag not found', () => {
            const xml = '<root><child/></root>';
            const result = insertBeforeTag(xml, '<notfound/>', '<inserted/>');
            expect(result).toBe(xml);
        });
    });

    describe('insertAfterTag', () => {
        it('should insert content after tag', () => {
            const xml = '<root><child/></root>';
            const result = insertAfterTag(xml, '<child/>', '<inserted/>');
            expect(result).toBe('<root><child/><inserted/></root>');
        });
    });

    describe('addRelationship', () => {
        it('should add relationship before closing tag', () => {
            const rels = '<Relationships><Relationship Id="rId1"/></Relationships>';
            const newRel = '<Relationship Id="rId2"/>';
            const result = addRelationship(rels, newRel);
            expect(result).toContain('rId2');
            expect(result).toContain('</Relationships>');
        });
    });

    describe('extractRelationshipIds', () => {
        it('should extract all relationship IDs', () => {
            const rels = '<Relationships><Relationship Id="rId1"/><Relationship Id="rId2"/></Relationships>';
            const result = extractRelationshipIds(rels);
            expect(result).toContain('rId1');
            expect(result).toContain('rId2');
        });

        it('should return empty array if no IDs', () => {
            const result = extractRelationshipIds('<Relationships></Relationships>');
            expect(result).toEqual([]);
        });
    });

    describe('generateUniqueRelId', () => {
        it('should generate unique ID not in existing list', () => {
            const existing = ['rId1', 'rId2', 'rId3'];
            const result = generateUniqueRelId(existing);
            expect(result).toBe('rId4');
        });

        it('should return rId1 for empty list', () => {
            const result = generateUniqueRelId([]);
            expect(result).toBe('rId1');
        });
    });

    describe('addContentType', () => {
        it('should add content type for new extension', () => {
            const xml = '<Types></Types>';
            const result = addContentType(xml, 'png', 'image/png');
            expect(result).toContain('Extension="png"');
            expect(result).toContain('ContentType="image/png"');
        });

        it('should not duplicate existing extension', () => {
            const xml = '<Types><Default Extension="png" ContentType="image/png"/></Types>';
            const result = addContentType(xml, 'png', 'image/png');
            expect((result.match(/Extension="png"/g) || []).length).toBe(1);
        });
    });

    describe('getImageContentType', () => {
        it('should return correct content type for known extensions', () => {
            expect(getImageContentType('png')).toBe('image/png');
            expect(getImageContentType('jpg')).toBe('image/jpeg');
            expect(getImageContentType('jpeg')).toBe('image/jpeg');
            expect(getImageContentType('gif')).toBe('image/gif');
        });

        it('should handle extension with dot', () => {
            expect(getImageContentType('.png')).toBe('image/png');
        });

        it('should default to png for unknown extensions', () => {
            expect(getImageContentType('unknown')).toBe('image/png');
        });
    });
});
