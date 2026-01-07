import { getDocumentProxy } from 'unpdf';

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
