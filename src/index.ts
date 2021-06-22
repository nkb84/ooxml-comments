// import { DocxService, PptxService, XlsxService, Service } from "./ooxml/service"

import { ThreadedComment } from "./ooxml/entity/comment"
import { OOXmlExtractor } from "./ooxml/extractor/ooxml_extractor"

const main = async (filePath: string) => {
  let service = OOXmlExtractor.from(filePath)
  const commentList: ThreadedComment[] = await service.getCommentList()
  console.log(JSON.stringify(commentList))
  console.log('Dump')
  service.dump()
}

main(process.argv[2])