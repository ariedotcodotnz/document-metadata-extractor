# Document Metadata Extractor

A TypeScript library for extracting metadata from various document types. This library provides a unified interface for extracting metadata from PDFs, images, Excel files, Word documents, and PowerPoint presentations.

## Overview

This library is built on top of various specialized libraries to extract metadata from different document formats. Each document type uses its underlying library to parse and extract relevant metadata:

- **PDF**: Built on top of `unpdf` for extracting PDF metadata and page counts
- **Images**: Built on top of `exiftool-vendored` for extracting EXIF and image metadata
- **Excel**: Built on top of `xlsx` for extracting spreadsheet metadata, sheet information, and document properties
- **DOCX/PPTX**: Built on top of `jszip` and `@xmldom/xmldom` for parsing Office Open XML documents and extracting metadata from core and application properties

## Installation

```bash
npm install @xcvzmoon/document-metadata-extractor
# or
pnpm add @xcvzmoon/document-metadata-extractor
# or
yarn add @xcvzmoon/document-metadata-extractor
# or
bun add @xcvzmoon/document-metadata-extractor
```

## Usage

```typescript
import { getMetadata } from '@xcvzmoon/document-metadata-extractor';
import { readFile } from 'fs/promises';

// Read a file as Buffer
const fileBuffer = await readFile('document.pdf');

// Extract metadata
const metadata = await getMetadata(fileBuffer, { target: 'pdf' });
console.log(metadata);
```

## Supported Document Types

### PDF

Extracts PDF metadata including title, author, subject, creator, producer, creation date, modification date, and page count.

```typescript
const metadata = await getMetadata(pdfBuffer, { target: 'pdf' });
// Returns: PdfMetadata with pages, title, author, subject, creator, producer, creationDate, modificationDate
```

### Images

Extracts EXIF and image metadata using ExifTool. Returns all available tags from the image file.

```typescript
const metadata = await getMetadata(imageBuffer, { target: 'image' });
// Returns: All ExifTool tags for the image
```

### Excel

Extracts spreadsheet metadata including sheet names, sheet count, row/column counts, author, last modified by, creation/modification dates, company, and file size.

```typescript
const metadata = await getMetadata(excelBuffer, { target: 'excel' });
// Returns: ExcelMetadata with sheets, sheetCount, rows, columns, author, lastModifiedBy, created, modified, company, fileSize
```

### DOCX

Extracts Word document metadata including title, subject, creator, keywords, description, last modified by, revision, creation/modification dates, category, company, page count, word count, character count, and file size.

```typescript
const metadata = await getMetadata(docxBuffer, { target: 'docx' });
// Returns: DocxMetadata with title, subject, creator, keywords, description, lastModifiedBy, revision, created, modified, category, company, pageCount, wordCount, characterCount, fileSize
```

### PPTX

Extracts PowerPoint presentation metadata using the same extraction method as DOCX files.

```typescript
const metadata = await getMetadata(pptxBuffer, { target: 'pptx' });
// Returns: DocxMetadata (same structure as DOCX)
```

## API

### `getMetadata(data: Buffer, options: { target: 'image' | 'pdf' | 'docx' | 'excel' | 'pptx' })`

Extracts metadata from a document buffer based on the specified target type.

**Parameters:**

- `data`: A Buffer containing the document file data
- `options.target`: The document type to extract metadata from

**Returns:**

- Promise resolving to the appropriate metadata type based on the target:
  - `PdfMetadata` for PDF files
  - ExifTool tags object for images
  - `ExcelMetadata` for Excel files
  - `DocxMetadata` for DOCX and PPTX files

## Type Definitions

The library exports TypeScript type definitions for all metadata types:

- `PdfMetadata`
- `ExcelMetadata`
- `DocxMetadata`

## License

ISC

## Author

Mon Albert Gamil - [GitHub](https://github.com/xcvzmoon/)
