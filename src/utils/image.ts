/**
 * Image processing utilities for DOCX
 */

import * as fs from 'fs';
import * as path from 'path';
import axios from 'axios';
import { ImageConfig } from '../types';
import { generateImageId } from './string';
import { generateImageRelationship, generateInlineImage } from './constants';

/**
 * Load image from file path
 */
export const loadImageFromPath = async (imagePath: string): Promise<Buffer> => {
    return fs.promises.readFile(imagePath);
};

/**
 * Download image from URL
 */
export const downloadImage = async (url: string): Promise<Buffer> => {
    const response = await axios({
        method: 'GET',
        url: url,
        responseType: 'arraybuffer'
    });
    return Buffer.from(response.data);
};

/**
 * Get image buffer from ImageConfig
 */
export const getImageBuffer = async (config: ImageConfig): Promise<Buffer> => {
    if (config.buffer) {
        return config.buffer;
    }

    if (config.path) {
        return loadImageFromPath(config.path);
    }

    if (config.url) {
        return downloadImage(config.url);
    }

    throw new Error(`No image source provided for placeholder: ${config.placeholder}`);
};

/**
 * Get image extension from path or URL
 */
export const getImageExtension = (config: ImageConfig): string => {
    if (config.path) {
        return path.extname(config.path).toLowerCase().replace('.', '') || 'png';
    }

    if (config.url) {
        const urlPath = new URL(config.url).pathname;
        return path.extname(urlPath).toLowerCase().replace('.', '') || 'png';
    }

    return 'png';
};

/**
 * Generate image file name for media folder
 */
export const generateImageFileName = (index: number, extension: string): string => {
    return `image${index + 1}.${extension}`;
};

/**
 * Prepare image for insertion into DOCX
 */
export interface PreparedImage {
    id: string;
    buffer: Buffer;
    fileName: string;
    extension: string;
    relationshipXml: string;
    inlineXml: string;
    placeholder: string;
}

/**
 * Prepare an image config for insertion
 */
export const prepareImage = async (
    config: ImageConfig,
    index: number,
    existingIds: string[]
): Promise<PreparedImage> => {
    const buffer = await getImageBuffer(config);
    const extension = getImageExtension(config);
    const fileName = generateImageFileName(index, extension);

    // Generate unique ID
    let id = config.id || generateImageId();
    let counter = 1;
    while (existingIds.includes(id)) {
        id = `rId${100 + index + counter}`;
        counter++;
    }

    const relationshipXml = generateImageRelationship(id, fileName);
    const inlineXml = generateInlineImage(
        id,
        config.width || 914400,
        config.height || config.width || 914400,
        `Image ${index + 1}`
    );

    return {
        id,
        buffer,
        fileName,
        extension,
        relationshipXml,
        inlineXml,
        placeholder: config.placeholder
    };
};

/**
 * Default image dimensions in EMUs (914400 EMUs = 1 inch)
 */
export const EMU_PER_INCH = 914400;
export const EMU_PER_CM = 360000;
export const EMU_PER_PIXEL = 9525; // At 96 DPI

/**
 * Convert pixels to EMUs
 */
export const pixelsToEmu = (pixels: number): number => {
    return Math.round(pixels * EMU_PER_PIXEL);
};

/**
 * Convert inches to EMUs
 */
export const inchesToEmu = (inches: number): number => {
    return Math.round(inches * EMU_PER_INCH);
};

/**
 * Convert centimeters to EMUs
 */
export const cmToEmu = (cm: number): number => {
    return Math.round(cm * EMU_PER_CM);
};
