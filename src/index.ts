import * as xlsx from 'xlsx';
import jszip from 'jszip';

import { DOMParser } from '@xmldom/xmldom';
import { getDocumentProxy } from 'unpdf';
import { ExifTool } from 'exiftool-vendored';
import { writeFile, unlink } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from 'path';
import { randomUUID } from 'crypto';

export type PdfMetadata = Awaited<
  ReturnType<Awaited<ReturnType<typeof getDocumentProxy>>['getMetadata']>
> & {
  pages: number;
};

export type ExcelMetadata = {
  sheets: string[];
  sheetCount: number;
  rows: number;
  columns: number;
  author?: string;
  lastModifiedBy?: string;
  created?: Date;
  modified?: Date;
  company?: string;
  fileSize: number;
};

export type DocxMetadata = {
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  revision?: string;
  created?: Date;
  modified?: Date;
  category?: string;
  company?: string;
  pageCount?: number;
  wordCount?: number;
  characterCount?: number;
  fileSize: number;
};

async function getPdfMetadata(data: Buffer) {
  const arrayBuffer = new Uint8Array(data);
  const pdf = await getDocumentProxy(arrayBuffer);
  const metadata = await pdf.getMetadata();
  const pages = pdf.numPages;
  return { ...metadata, pages } as PdfMetadata;
}

async function getImageMetadata(data: Buffer) {
  await using et = new ExifTool();
  const path = join(tmpdir(), `exif-${randomUUID()}`);

  try {
    await writeFile(path, data);
    const tags = await et.read(path);
    return tags;
  } finally {
    await unlink(path);
  }
}

function getExcelMetadata(data: Buffer) {
  const workbook = xlsx.read(data, {
    type: 'buffer',
    cellDates: true,
  });

  let rows = 0;
  let columns = 0;

  const sheets = workbook.SheetNames;
  for (const name of sheets) {
    const sheet = workbook.Sheets[name];
    const range = xlsx.utils.decode_range(sheet['!ref'] || 'A1');

    rows += range.e.r + 1;
    columns += Math.max(columns, range.e.c + 1);
  }

  const props = workbook.Props || {};

  return {
    sheets,
    sheetCount: sheets.length,
    rows,
    columns,
    author: props.Author,
    lastModifiedBy: props.LastAuthor,
    created: props.CreatedDate,
    modified: props.ModifiedDate,
    company: props.Company,
    fileSize: data.length,
  };
}

async function getDocxMetadata(data: Buffer) {
  const zip = await jszip.loadAsync(data);
  const core = await zip.file('docProps/core.xml')?.async('string');
  const app = await zip.file('docProps/app.xml')?.async('string');
  const metadata: DocxMetadata = { fileSize: data.length };

  if (core) {
    const parser = new DOMParser();
    const document = parser.parseFromString(core, 'text/xml');

    metadata.title = getTextContent(document, 'dc:title');
    metadata.subject = getTextContent(document, 'dc:subject');
    metadata.creator = getTextContent(document, 'dc:creator');
    metadata.keywords = getTextContent(document, 'cp:keywords');
    metadata.description = getTextContent(document, 'dc:description');
    metadata.lastModifiedBy = getTextContent(document, 'cp:lastModifiedBy');
    metadata.revision = getTextContent(document, 'cp:revision');
    metadata.category = getTextContent(document, 'cp:category');

    const created = getTextContent(document, 'dcterms:created');
    const modified = getTextContent(document, 'dcterms:modified');

    if (created) metadata.created = new Date(created);
    if (modified) metadata.modified = new Date(modified);
  }

  if (app) {
    const parser = new DOMParser();
    const document = parser.parseFromString(app, 'text/xml');

    metadata.company = getTextContent(document, 'Company');

    const pages = getTextContent(document, 'Pages');
    const words = getTextContent(document, 'Words');
    const characters = getTextContent(document, 'Characters');

    if (pages) metadata.pageCount = parseInt(pages);
    if (words) metadata.wordCount = parseInt(words);
    if (characters) metadata.characterCount = parseInt(characters);
  }

  return metadata;
}

function getTextContent(document: Document, tagName: string): string | undefined {
  const elements = document.getElementsByTagName(tagName);
  return elements.length > 0 ? elements[0].textContent || undefined : undefined;
}

export async function getMetadata(
  data: Buffer,
  options: { target: 'image' | 'pdf' | 'docx' | 'excel' | 'pptx' },
) {
  switch (options.target) {
    case 'pdf':
      const pdfMetadata = await getPdfMetadata(data);
      return pdfMetadata;

    case 'image':
      const imageMetadata = await getImageMetadata(data);
      return imageMetadata;

    case 'excel':
      const excelMetadata = getExcelMetadata(data);
      return excelMetadata;

    case 'docx':
      const docxMetadata = await getDocxMetadata(data);
      return docxMetadata;

    case 'pptx':
      const pptxMetadata = await getDocxMetadata(data);
      return pptxMetadata;
  }
}
