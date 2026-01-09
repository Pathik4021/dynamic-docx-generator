/**
 * Unit tests for string utilities
 */

import {
    escapeRegExp,
    replaceAll,
    replacePlaceholder,
    escapeForXml,
    unescapeXml,
    fixDoubleEscaping,
    stripHtml,
    generateImageId,
    extractPlaceholders
} from '../../src/utils/string';

describe('String Utilities', () => {
    describe('escapeRegExp', () => {
        it('should escape special regex characters', () => {
            expect(escapeRegExp('test.*+?')).toBe('test\\.\\*\\+\\?');
            expect(escapeRegExp('{{placeholder}}')).toBe('\\{\\{placeholder\\}\\}');
        });

        it('should handle empty string', () => {
            expect(escapeRegExp('')).toBe('');
        });
    });

    describe('replaceAll', () => {
        it('should replace all occurrences', () => {
            expect(replaceAll('hello world hello', 'hello', 'hi')).toBe('hi world hi');
        });

        it('should handle special characters', () => {
            expect(replaceAll('{{name}} and {{name}}', '{{name}}', 'John')).toBe('John and John');
        });

        it('should return original if not a string', () => {
            expect(replaceAll(null as any, 'a', 'b')).toBe(null);
        });
    });

    describe('replacePlaceholder', () => {
        it('should replace {{placeholder}} format', () => {
            const result = replacePlaceholder('Hello {{name}}!', 'name', 'World');
            expect(result).toBe('Hello World!');
        });

        it('should replace <w:t>placeholder</w:t> format', () => {
            const result = replacePlaceholder('<w:t>name</w:t>', 'name', 'John');
            expect(result).toBe('<w:t>John</w:t>');
        });

        it('should escape XML special characters in value', () => {
            const result = replacePlaceholder('{{content}}', 'content', 'A & B');
            expect(result).toBe('A &amp; B');
        });
    });

    describe('escapeForXml', () => {
        it('should escape XML special characters', () => {
            expect(escapeForXml('A & B')).toBe('A &amp; B');
            expect(escapeForXml('<tag>')).toBe('&lt;tag&gt;');
            expect(escapeForXml('"quoted"')).toBe('&quot;quoted&quot;');
        });

        it('should handle empty or null input', () => {
            expect(escapeForXml('')).toBe('');
            expect(escapeForXml(null as any)).toBe('');
        });
    });

    describe('unescapeXml', () => {
        it('should unescape XML entities', () => {
            expect(unescapeXml('A &amp; B')).toBe('A & B');
            expect(unescapeXml('&lt;tag&gt;')).toBe('<tag>');
        });
    });

    describe('fixDoubleEscaping', () => {
        it('should fix double-escaped ampersands', () => {
            expect(fixDoubleEscaping('A &amp;amp; B')).toBe('A &amp; B');
        });

        it('should fix double-escaped less than', () => {
            expect(fixDoubleEscaping('&amp;lt;tag&amp;gt;')).toBe('&lt;tag&gt;');
        });
    });

    describe('stripHtml', () => {
        it('should remove HTML tags', () => {
            expect(stripHtml('<p>Hello <b>World</b></p>')).toBe('Hello World');
        });

        it('should handle empty input', () => {
            expect(stripHtml('')).toBe('');
        });
    });

    describe('generateImageId', () => {
        it('should generate a string starting with rId', () => {
            const id = generateImageId();
            expect(id.startsWith('rId')).toBe(true);
        });

        it('should generate unique IDs', () => {
            const ids = new Set<string>();
            for (let i = 0; i < 100; i++) {
                ids.add(generateImageId());
            }
            expect(ids.size).toBeGreaterThan(90); // Allow some collisions
        });
    });

    describe('extractPlaceholders', () => {
        it('should extract {{}} style placeholders', () => {
            const content = '{{name}} is {{age}} years old';
            const result = extractPlaceholders(content);
            expect(result).toContain('name');
            expect(result).toContain('age');
        });

        it('should return empty array if no placeholders', () => {
            expect(extractPlaceholders('No placeholders here')).toEqual([]);
        });
    });
});
