/**
 * Configuration options for DocxGenerator
 */
export interface GeneratorOptions {
    /** Custom temporary directory for processing */
    tempDir?: string;
}

/**
 * Configuration for an image to be inserted into the document
 */
export interface ImageConfig {
    /** The placeholder text in the template to replace (e.g., "{{logo}}") */
    placeholder: string;
    /** Path to the image file */
    path?: string;
    /** Image data as a Buffer */
    buffer?: Buffer;
    /** URL to download the image from */
    url?: string;
    /** Width of the image in EMUs (914400 EMUs = 1 inch). Default: 914400 */
    width?: number;
    /** Height of the image in EMUs (914400 EMUs = 1 inch). Default: auto */
    height?: number;
    /** Unique ID for the image relationship */
    id?: string;
}

/**
 * Table styling options
 */
export interface TableStyle {
    /** Header row background color (hex without #) */
    headerBgColor?: string;
    /** Header text color (hex without #) */
    headerTextColor?: string;
    /** Border color (hex without #) */
    borderColor?: string;
    /** Font size in half-points (24 = 12pt) */
    fontSize?: number;
    /** Font family */
    fontFamily?: string;
    /** Header row height in twips (1/20th of a point). Default: 400 */
    headerHeight?: number;
    /** Data row height in twips (1/20th of a point). Default: 350 */
    rowHeight?: number;
    /** Table alignment ('center' | 'left' | 'right'). Default: 'center' */
    tableAlign?: 'center' | 'left' | 'right';
    /** Cell padding in twips. Default: 100 */
    cellPadding?: number;
    /** Border size in eighths of a point (4 = 1/2pt). Default: 4 */
    borderSize?: number;
}

/**
 * Configuration for a dynamic table
 */
export interface TableConfig {
    /** The placeholder text in the template to replace */
    placeholder: string;
    /** Array of header column names */
    headers: TableHeader[];
    /** 2D array of row data */
    rows: string[][];
    /** Optional styling */
    style?: TableStyle;
}

/**
 * Table header configuration
 */
export interface TableHeader {
    /** Header display name */
    name: string;
    /** Data key for mapping */
    key: string;
    /** Column width in twips (1440 twips = 1 inch) */
    width?: number;
}

/**
 * Placeholder data for simple text replacement
 */
export type PlaceholderData = Record<string, string>;

/**
 * Result of document generation
 */
export interface GenerationResult {
    /** Whether generation was successful */
    success: boolean;
    /** Output file path or buffer */
    output?: string | Buffer;
    /** Error message if failed */
    error?: string;
}
