import { generateTable, DEFAULT_TABLE_STYLE } from '../../src/utils/constants';

describe('Custom Table Styling', () => {
    const headers = [
        { name: 'Header 1', width: 1000 },
        { name: 'Header 2', width: 2000 }
    ];
    const rows = [
        ['Row 1 Col 1', 'Row 1 Col 2'],
        ['Row 2 Col 1', 'Row 2 Col 2']
    ];

    it('should use default styles when no custom styles are provided', () => {
        const xml = generateTable(headers, rows);

        expect(xml).toContain('<w:trHeight w:val="400" />'); // Default header height
        expect(xml).toContain('<w:trHeight w:val="350" />'); // Default row height
        expect(xml).toContain('<w:jc w:val="center" />'); // Default table align
        expect(xml).toContain('<w:top w:val="single" w:sz="4"'); // Default border size
        expect(xml).toContain('<w:top w:w="100" w:type="dxa"/>'); // Default padding
    });

    it('should apply custom row and header heights', () => {
        const style = {
            ...DEFAULT_TABLE_STYLE,
            headerHeight: 600,
            rowHeight: 500
        };
        const xml = generateTable(headers, rows, style);

        expect(xml).toContain('<w:trHeight w:val="600" />');
        expect(xml).toContain('<w:trHeight w:val="500" />');
    });

    it('should apply custom table alignment', () => {
        const style = {
            ...DEFAULT_TABLE_STYLE,
            tableAlign: 'left' as const
        };
        const xml = generateTable(headers, rows, style);

        expect(xml).toContain('<w:jc w:val="left" />');
    });

    it('should apply custom cell padding', () => {
        const style = {
            ...DEFAULT_TABLE_STYLE,
            cellPadding: 200
        };
        const xml = generateTable(headers, rows, style);

        // Check header padding
        expect(xml).toContain('<w:top w:w="200" w:type="dxa"/>');
        expect(xml).toContain('<w:left w:w="200" w:type="dxa"/>');

        // Check row padding count (headers + rows) * cols * 4 sides
        // Just ensuring it's present for now
        expect(xml).toContain('<w:bottom w:w="200" w:type="dxa"/>');
    });

    it('should apply custom border size', () => {
        const style = {
            ...DEFAULT_TABLE_STYLE,
            borderSize: 8
        };
        const xml = generateTable(headers, rows, style);

        expect(xml).toContain('<w:top w:val="single" w:sz="8"');
        expect(xml).toContain('<w:left w:val="single" w:sz="8"');
    });
});
