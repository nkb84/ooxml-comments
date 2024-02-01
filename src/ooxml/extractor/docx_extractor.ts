import { ThreadedComment } from "../entity/comment";
import { ContentType, ContentTypes } from "../entity/content_type";
import Person from "../entity/person";
import { Relation } from "../entity/relation";
import { Parser } from "../utils/parser";
import { CommentList } from "./extractor";
import { BaseExtractor } from "./base_extractor";

const path = require('path')

export class DocxExtractor extends BaseExtractor {
  private static ONS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  private static W14NS = "http://schemas.microsoft.com/office/word/2010/wordml"
  private static W15NS = "http://schemas.microsoft.com/office/word/2012/wordml"
  private static WPNS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  private document?: string // workbook name
  private commentPart?: string // comment part
  private commentExtenedPart?: string // extended part
  private threadedComments: ThreadedComment[] = []
  private comments: CommentList = {}
  private commentDetails: {[key: string]: string} = {}
  private headingMap: {[key: string]: string} = {} // Map between paraId with heading just above

  private rootRelationLoaded: boolean = false
  private documentRelationLoaded: boolean = false
  private commentLoaded: boolean = false

  constructor (path: string) {
    super(path)
  }

  adjustCommentList(): ThreadedComment[] {
    return this.threadedComments
  }

  loadContent(partName: string, data: string) : boolean {
    const contentType = this.getContentType(partName)
    if (contentType === undefined) {
      throw new Error(`Part ${partName} cannot be identified the content type`)
    }

    switch (contentType) {
      case ContentType.DOCUMENT: return this.loadCommentLocationFromMain(data); break;
      case ContentType.WORD_PEOPLE: return this.loadPersons(data); break;
      case ContentType.WORD_COMMENT: return this.loadComments(partName, data); break;
      case ContentType.WORD_COMMENT_EXTENDED: return this.loadCommentExtendeds(partName, data); break;
      case ContentType.RELATION: return this.loadRelations(partName, data); break;
      default:
        break;
    }

    return true
  }

  private createHeadingMap(doc: Document): void {
    const bodies = doc.getElementsByTagNameNS(DocxExtractor.ONS, 'body')
    if (bodies.length <= 0) {
      throw new Error(`body is not found in document`)
    }

    let lastHeading
    const children = bodies[0].childNodes
    for (let i = 0; i < children.length; i++) {
      const element = children[i] as Element
      if (element.localName !== 'p') {
        console.warn(`Element ${i} is not p`)
      } else {
        const paraId = element.getAttributeNS(DocxExtractor.W14NS, 'paraId')
        if (!paraId) {
          // throw new Error(`Element ${i} has no paraId`)
          continue
        }
        const pPrs = element.getElementsByTagNameNS(DocxExtractor.ONS, 'pStyle')
        if (pPrs.length > 0) {
          const style = pPrs[0].getAttributeNS(DocxExtractor.ONS, 'val')
          if (style?.startsWith('Heading')) {
            lastHeading = style + ': ' + element.textContent
          }
        }

        this.headingMap[paraId] = lastHeading || ""
      }
    }
  }

  private getNearestParagraph(element: Element): string | null {
    let p = element
    while (true) {
      while (p !== null && p.localName !== 'p') {
        if (p.parentNode !== null && (p.parentNode as Element).localName === 'body') {
          // Finish, move to previous sibling
          p = p.previousSibling as Element
        } else {
          p = p.parentNode as Element
        }
      }

      if (p === null) {
        // Not found
        return null
      }
      const paraId = p.getAttributeNS(DocxExtractor.W14NS, 'paraId')
      if (paraId && this.headingMap.hasOwnProperty(paraId)) {
        return paraId
      }

      // Not the p we need
      p = p.parentNode as Element
    }
  }

  private loadCommentLocationFromMain(content: string): boolean {
    const parser = Parser.getInstance()
    try {
      const doc = parser.parseFromString(content)
      // Create heading map
      this.createHeadingMap(doc)

      // Process all comment ranges
      const ranges = doc.getElementsByTagNameNS(DocxExtractor.ONS, 'commentRangeStart')
      for (let i = 0; i < ranges.length; i++) {
        const rangeStart = ranges[i]
        const id = rangeStart.getAttributeNS(DocxExtractor.ONS, 'id') || ""
        let commentNode = rangeStart.nextSibling as Element
        while (commentNode.localName !== 'r') {
          commentNode = commentNode.nextSibling as Element
        }
        let commentTxt = commentNode.textContent
        if (!commentTxt) {
          const drawings = commentNode.getElementsByTagNameNS(DocxExtractor.ONS, 'drawing')
          if (drawings.length) {
            const docPtrs = drawings[0].getElementsByTagNameNS(DocxExtractor.WPNS, 'docPr')
            if (docPtrs.length) {
              commentTxt = docPtrs[0].getAttribute('name')
            }
          }
        }

        const parentParaId = this.getNearestParagraph(rangeStart)
        if (!parentParaId) {
          throw new Error(`CommentRangeStrt id = ${id} not found valid parent p`)
        }
        this.commentDetails[id] = this.headingMap[parentParaId] + ', selected: "' + commentTxt + '"'
      }
    } catch (e) {
      console.log(e)
    }

    return true
  }

  private loadRelations(partName: string, content: string): boolean {
    const wbRelsPath = Relation.getRelsPath(this.document)
    const parser = Parser.getInstance()
    try {
      const doc = parser.parseFromString(content)
      const fname = partName.split('/').pop()?.toLocaleLowerCase()
      if (fname === undefined) {
        throw new Error(`Part ${partName} error: cannot get file name from it`)
      } else if (fname === '.rels' && !this.rootRelationLoaded) {
        // get workbook name
        const relations = doc.getElementsByTagName('Relationship')
        for (let i = 0; i < relations.length; i++) {
          if (relations[i].getAttribute('Type') === Relation.OFFICE_DOCUMENT) {
            this.document = (relations[i].getAttribute('Target') || "").replace('\\', '/')
          }
        }
        this.rootRelationLoaded = true
      } else if (partName === wbRelsPath && !this.documentRelationLoaded) {
        // Workbook relations, get all sheets
        const relations = doc.getElementsByTagName('Relationship')
        for (let i = 0; i < relations.length; i++) {
          const stype = relations[i].getAttribute('Type') || ""
          if ([Relation.COMMENT, Relation.COMMENT_EXTENDED].includes(stype)) {
            const starget = (relations[i].getAttribute('Target') || "").replace('\\', '/')
            // Get sheet name
            const sname = starget.split('/').pop()?.split('.').slice(0, -1).join('.')
            // Workbook folder
            const paths = this.document!.split('/')
            paths.pop()
            const spartName = path.posix.normalize(paths.join('/') + '/' + starget)
            if (sname === undefined || spartName === undefined) {
              throw new Error(`Sheet name or target name is underfine for ${starget}`)
            }
            if (stype === Relation.COMMENT) {
              this.commentPart = spartName
            } else {
              this.commentExtenedPart = spartName
            }
          }
        }

        this.documentRelationLoaded = true
      } else {
        return false
      }
    } catch (e) {
      console.log(e)
    }

    return true
  }

  updateThreadComments() {
    for (let id in this.comments) {
      const c = this.comments[id]
      if (!this.commentDetails.hasOwnProperty(c.ref)) {
        console.warn(`Has not comment detail for ${c.ref}`)
      } else {
        c.location = this.commentDetails[c.ref]
      }
      if (c.parentId === "") {
        // This is a root object
        this.threadedComments.push(c)
      } else {
        // Get comment by id
        const parent = this.getCommentById(c.parentId)
        if (parent === undefined) {
          throw new Error(`Comment ${c.parentId} is not found`)
        }
        if (parent.children === undefined) {
          parent.children = [c]
        } else {
          parent.children.push(c)
        }
      }
    }
  }

  private getCommentById(id: string): ThreadedComment | undefined {
    if (id in this.comments) {
      return this.comments[id]
    }
    return undefined
  }

  private loadComments(partName: string, content: string): boolean {
    if (this.commentLoaded) {
      throw new Error('Comment already loaded')
    }

    const parser = Parser.getInstance()
    try {
      const doc = parser.parseFromString(content)
      const commentList = doc.getElementsByTagNameNS(DocxExtractor.ONS, 'comment')
      for (let i = 0; i < commentList.length; i++) {
        const element = commentList[i]
        const paraElement = element.firstChild as Element
        if (paraElement === null) {
          throw new Error('Comment node must have a child')
        }
        const id = paraElement.getAttributeNS(DocxExtractor.W14NS, 'paraId') || ""
        this.comments[id] = new ThreadedComment(
                      id,
                      element.getAttributeNS(DocxExtractor.ONS, 'id') || "",
                      element.getAttributeNS(DocxExtractor.ONS, 'date') || "",
                      element.getAttributeNS(DocxExtractor.ONS, 'author') || "",
                      "", // parent id
                      false, // done
                      element.textContent || ""
                    )
          this.comments[id].partName = partName
          this.comments[id].location = this.document
      }

      // Thread comment need to be loaded last
      // this.updateThreadComments()
    } catch (e) {
      console.log(e)
    }

    this.commentLoaded = true
    return true
  }

  private loadCommentExtendeds(partName: string, content: string): boolean {
    if (!this.commentLoaded) {
      return false
    }

    const parser = Parser.getInstance()
    try {
      const doc = parser.parseFromString(content)
      const commentList = doc.getElementsByTagNameNS(DocxExtractor.W15NS, 'commentEx')
      for (let i = 0; i < commentList.length; i++) {
        const element = commentList[i]
        const id = element.getAttributeNS(DocxExtractor.W15NS, 'paraId') || ""
        if (!this.getCommentById(id)) {
          throw new Error(`Comment with id ${id} not found`)
        }

        const comment = this.getCommentById(id)
        comment!.done = element.getAttributeNS(DocxExtractor.W15NS, 'done') === '1'
        comment!.parentId = element.getAttributeNS(DocxExtractor.W15NS, 'paraIdParent') || ""
      }

      // Thread comment need to be loaded last
      // this.updateThreadComments()
    } catch (e) {
      console.log(e)
    }

    return true
  }

  private loadPersons(content: string): boolean {
    const parser = Parser.getInstance()

    try {
      const doc = parser.parseFromString(content)
      const personList = doc.getElementsByTagNameNS(DocxExtractor.W15NS, 'person')
      for (let i = 0; i < personList.length; i++) {
        const element = personList[i]
        const presenceElement = element.firstChild as Element
        if (presenceElement === null) {
          throw new Error('Person node must has a child')
        }
        const id = presenceElement.getAttributeNS(DocxExtractor.W15NS, 'userId') || ""
        this.persons[id] = new Person(
                      element.getAttributeNS(DocxExtractor.W15NS, 'author') || "",
                      id,
                      presenceElement.getAttributeNS(DocxExtractor.W15NS, 'providerId') || ""
                  )
      }
    } catch (e) {
      console.log(e)
    }

    return true
  }

  dumpComments() {
    this.threadedComments.forEach(c => {
      this.dumpComment(1, c)
    })
  }
}