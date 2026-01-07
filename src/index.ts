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
