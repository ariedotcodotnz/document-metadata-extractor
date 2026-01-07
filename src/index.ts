import * as xlsx from 'xlsx';

import { getDocumentProxy } from 'unpdf';
import { ExifTool } from 'exiftool-vendored';
import { writeFile, unlink } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from 'path';
import { randomUUID } from 'crypto';

export type PdfMetadata = Promise<
  Awaited<ReturnType<Awaited<ReturnType<typeof getDocumentProxy>>['getMetadata']>> & {
    pages: number;
  }
>;

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

export async function getPdfMetadata(data: ArrayBuffer): PdfMetadata {
  const arrayBuffer = new Uint8Array(data);
  const pdf = await getDocumentProxy(arrayBuffer);
  const metadata = await pdf.getMetadata();
  const pages = pdf.numPages;

  return { ...metadata, pages };
}

export async function getImageMetadata(data: Buffer) {
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

export function getExcelMetadata(data: Buffer) {
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
