/**
 * XML template constants for Office Open XML (OOXML) DOCX format
 */

/**
 * Generate XML for image relationship entry in .rels file
 */
export const generateImageRelationship = (imageId: string, imageName: string): string => {
    return `<Relationship Id="${imageId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${imageName}" />`;
};

/**
 * Generate XML for inline image in document
 */
export const generateInlineImage = (
    imageId: string,
    width: number = 914400,
    height: number = 914400,
    name: string = 'Picture'
): string => {
    return `<w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="${width}" cy="${height}" />
      <wp:effectExtent l="0" t="0" r="0" b="0" />
      <wp:docPr id="1" name="${name}" />
      <wp:cNvGraphicFramePr>
        <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" />
      </wp:cNvGraphicFramePr>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:nvPicPr>
              <pic:cNvPr id="1" name="${name}" />
              <pic:cNvPicPr />
            </pic:nvPicPr>
            <pic:blipFill>
              <a:blip r:embed="${imageId}" cstate="print" />
              <a:stretch>
                <a:fillRect />
              </a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0" />
                <a:ext cx="${width}" cy="${height}" />
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst />
              </a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;
};

/**
 * Default table style values
 */
export const DEFAULT_TABLE_STYLE = {
    headerBgColor: 'C9DAF8',
    headerTextColor: '000000',
    borderColor: 'auto',
    fontSize: 24, // 12pt
    fontFamily: 'Times New Roman'
};

/**
 * Generate table header row XML
 */
export const generateTableHeaderRow = (
    headers: { name: string; width?: number }[],
    style: typeof DEFAULT_TABLE_STYLE = DEFAULT_TABLE_STYLE
): string => {
    const cells = headers.map((header, index) => `
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="${header.width || 2000}" w:type="dxa" />
        <w:shd w:val="clear" w:color="auto" w:fill="${style.headerBgColor}" />
        <w:vAlign w:val="center" />
      </w:tcPr>
      <w:p>
        <w:pPr>
          <w:jc w:val="center" />
          <w:rPr>
            <w:rFonts w:ascii="${style.fontFamily}" w:hAnsi="${style.fontFamily}" />
            <w:b />
            <w:sz w:val="${style.fontSize}" />
            <w:szCs w:val="${style.fontSize}" />
            <w:color w:val="${style.headerTextColor}" />
          </w:rPr>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="${style.fontFamily}" w:hAnsi="${style.fontFamily}" />
            <w:b />
            <w:sz w:val="${style.fontSize}" />
            <w:szCs w:val="${style.fontSize}" />
            <w:color w:val="${style.headerTextColor}" />
          </w:rPr>
          <w:t>${escapeXml(header.name)}</w:t>
        </w:r>
      </w:p>
    </w:tc>
  `).join('');

    return `
    <w:tr>
      <w:trPr>
        <w:trHeight w:val="400" />
      </w:trPr>
      ${cells}
    </w:tr>
  `;
};

/**
 * Generate a single table data row XML
 */
export const generateTableDataRow = (
    cells: string[],
    widths: number[],
    style: typeof DEFAULT_TABLE_STYLE = DEFAULT_TABLE_STYLE
): string => {
    const cellsXml = cells.map((cell, index) => `
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="${widths[index] || 2000}" w:type="dxa" />
        <w:vAlign w:val="center" />
      </w:tcPr>
      <w:p>
        <w:pPr>
          <w:rPr>
            <w:rFonts w:ascii="${style.fontFamily}" w:hAnsi="${style.fontFamily}" />
            <w:sz w:val="${style.fontSize}" />
            <w:szCs w:val="${style.fontSize}" />
          </w:rPr>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="${style.fontFamily}" w:hAnsi="${style.fontFamily}" />
            <w:sz w:val="${style.fontSize}" />
            <w:szCs w:val="${style.fontSize}" />
          </w:rPr>
          <w:t>${escapeXml(cell)}</w:t>
        </w:r>
      </w:p>
    </w:tc>
  `).join('');

    return `
    <w:tr>
      <w:trPr>
        <w:trHeight w:val="350" />
      </w:trPr>
      ${cellsXml}
    </w:tr>
  `;
};

/**
 * Generate complete table XML
 */
export const generateTable = (
    headers: { name: string; width?: number }[],
    rows: string[][],
    style: typeof DEFAULT_TABLE_STYLE = DEFAULT_TABLE_STYLE
): string => {
    const widths = headers.map(h => h.width || 2000);
    const gridCols = headers.map(h => `<w:gridCol w:w="${h.width || 2000}" />`).join('');

    const headerRow = generateTableHeaderRow(headers, style);
    const dataRows = rows.map(row => generateTableDataRow(row, widths, style)).join('');

    return `
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid" />
        <w:tblW w:w="5000" w:type="pct" />
        <w:jc w:val="center" />
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="${style.borderColor}" />
          <w:left w:val="single" w:sz="4" w:space="0" w:color="${style.borderColor}" />
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="${style.borderColor}" />
          <w:right w:val="single" w:sz="4" w:space="0" w:color="${style.borderColor}" />
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="${style.borderColor}" />
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="${style.borderColor}" />
        </w:tblBorders>
        <w:tblLayout w:type="fixed" />
      </w:tblPr>
      <w:tblGrid>
        ${gridCols}
      </w:tblGrid>
      ${headerRow}
      ${dataRows}
    </w:tbl>
  `;
};

/**
 * Escape special XML characters
 */
export const escapeXml = (text: string): string => {
    if (!text) return '';
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
};
