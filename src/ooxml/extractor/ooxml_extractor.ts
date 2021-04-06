import { DocxExtractor } from "./docx_extractor";
import { Extractor } from "./extractor";
import { PptxExtractor } from "./pptx_extractor";
import { XlsxExtractor } from "./xlsx_extractor";

export class OOXmlExtractor {
  public static from(filePath: string): Extractor {
    const ext = filePath.split('.').pop()?.toLocaleLowerCase()
    let service: Extractor
    switch (ext) {
      case 'xlsx':
        service = new XlsxExtractor(filePath)
        break
      case 'docx':
        service = new DocxExtractor(filePath)
        break
      case 'pptx':
        service = new PptxExtractor(filePath)
        break
      default:
        throw new Error(`File ${filePath} with extension ${ext} is unrecognizable`)
        break
    }
    return service
  }
}