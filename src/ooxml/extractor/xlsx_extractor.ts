import { ThreadedComment } from "../entity/comment";
import { ContentType } from "../entity/content_type";
import Person from "../entity/person";
import { Relation } from "../entity/relation";
import { Container } from "../entity/container";
import { Parser } from "../utils/parser";
import { CommentList, ContainerList } from "./extractor";
import { BaseExtractor } from "./base_extractor";

const path = require('path')

export class XlsxExtractor extends BaseExtractor {
  private workbook?: string // workbook name
  private sheets: ContainerList = {}  // map from sheet name to sheet object
  private comments: CommentList = {}
  private commentPartNameMap: {[key: string]: string} = {}  // map from comment partname to sheet name

  private rootRelationLoaded: boolean = false
  private workbookRelationLoaded: boolean = false

  constructor (path: string) {
    super(path)
  }

  adjustCommentList(): ThreadedComment[] {
    const result: ThreadedComment[] = []
    Object.values(this.sheets).forEach(s => {
      if (s.threadedComments.length > 0) {
        result.push(...s.threadedComments)
      }
    })
    return result
  }

  private isRootRelationLoaded(): boolean {
    return this.rootRelationLoaded
  }

  private isWorkbookRelationLoaded(): boolean {
    return this.workbookRelationLoaded
  }

  loadContent(partName: string, data: string) : boolean {
    const contentType = this.getContentType(partName)
    if (contentType === undefined) {
      throw new Error(`Part ${partName} cannot be identified the content type`)
    }

    switch (contentType) {
      case ContentType.EXCEL_PERSON: return this.loadPersons(data); break;
      case ContentType.EXCEL_THREADED_COMMENT: return this.loadThreadedComments(partName, data); break;
      case ContentType.EXCEL_COMMENT: return this.loadComments(partName, data); break;
      case ContentType.RELATION: return this.loadRelations(partName, data); break;
      default:
        break;
    }

    return true
  }

  private loadRelations(partName: string, content: string): boolean {
    const wbRelsPath = Relation.getRelsPath(this.workbook)
    const sheetReslPaths: {[key: string]: string} = {}
    Object.values(this.sheets).map(s => {
      sheetReslPaths[s.getRelsPath() || ""] = s.name
    })

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
            this.workbook = (relations[i].getAttribute('Target') || "").replace('\\', '/')
          }
        }
        this.rootRelationLoaded = true
      } else if (partName === wbRelsPath && !this.workbookRelationLoaded) {
        // Workbook relations, get all sheets
        const relations = doc.getElementsByTagName('Relationship')
        for (let i = 0; i < relations.length; i++) {
          if (relations[i].getAttribute('Type') === Relation.WORKSHEET) {
            const starget = (relations[i].getAttribute('Target') || "").replace('\\', '/')
            // Get sheet name
            const sname = starget.split('/').pop()?.split('.').slice(0, -1).join('.')
            // Workbook folder
            const paths = this.workbook!.split('/')
            paths.pop()
            const spartName = path.posix.normalize(paths.join('/') + '/' + starget)
            if (sname === undefined || spartName === undefined) {
              throw new Error(`Sheet name or target name is underfine for ${starget}`)
            }
            this.sheets[sname] = new Container(sname, spartName)
          }
        }

        this.workbookRelationLoaded = true
      } else if (partName in sheetReslPaths && this.rootRelationLoaded && this.workbookRelationLoaded) {
        // sheet
        const relations = doc.getElementsByTagName('Relationship')
        for (let i = 0; i < relations.length; i++) {
          if ([Relation.THREADED_COMMENT, Relation.COMMENT].includes(relations[i].getAttribute('Type') || "")) {
            const ttarget = (relations[i].getAttribute('Target') || "").replace('\\', '/')
            // sheet folder
            const paths = partName.split('/')
            paths.pop()
            paths.pop()
            // Comment partname
            const cpartName = path.posix.normalize(paths.join('/') + '/' + ttarget)
            // Update map
            this.commentPartNameMap[cpartName] = sheetReslPaths[partName]
          }
        }
      } else {
        return false
      }
    } catch (e) {
      console.log(e)
    }

    return true
  }

  private getSheetNameFromCommentPart(partName?: string): string | undefined {
    if (partName !== undefined && partName in this.commentPartNameMap) {
      return this.commentPartNameMap[partName]
    }

    return undefined
  }

  updateThreadComments() {
    for (let id in this.comments) {
      const c = this.comments[id]
      const sname = this.getSheetNameFromCommentPart(c.partName)
      if (sname === undefined) {
        throw new Error(`Sheet is empty in comment ${c.id}`)
      }

      // Update location for comment
      const sheet = this.sheets[sname]
      c.location = sname + ': ' + c.ref
      if (c.parentId === "") {
        // This is a root object
        sheet.threadedComments.push(c)
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
    return this.comments[id]
  }

  private loadComments(partName: string, content: string): boolean {
    const parser = Parser.getInstance()
    const authors: string[] = []

    try {
      const doc = parser.parseFromString(content)

      // Get author list
      const authorList = doc.getElementsByTagName('author')
      for (let i = 0; i < authorList.length; i++) {
        authors.push(authorList[i].textContent || "")
      }

      // Get comment list
      const commentList = doc.getElementsByTagName('comment')
      const date = new Date()
      const id0 = (new Date()).getTime()
      for (let i = 0; i < commentList.length; i++) {
        const element = commentList[i]
        // No need id, since not threaded
        const id = (id0 + i).toString()
        const authorIdx = parseInt(element.getAttribute('authorId') || "")
        this.comments[id] = new ThreadedComment(
                      id,
                      element.getAttribute('ref') || "",
                      date.toLocaleString(), // no date time, get current
                      authorIdx < authors.length ? authors[authorIdx] : "",
                      "", // no parent
                      false,
                      element.textContent || ""
                    )
          this.comments[id].partName = partName
      }

      // Thread comment need to be loaded last
      // this.updateThreadComments()
    } catch (e) {
      console.log(e)
    }

    return true
  }

  private loadThreadedComments(partName: string, content: string): boolean {
    const parser = Parser.getInstance()

    try {
      const doc = parser.parseFromString(content)
      const commentList = doc.getElementsByTagName('threadedComment')
      for (let i = 0; i < commentList.length; i++) {
        const element = commentList[i]
        const id = element.getAttribute('id') || ""
        this.comments[id] = new ThreadedComment(
                      element.getAttribute('id') || "",
                      element.getAttribute('ref') || "",
                      element.getAttribute('dT') || "",
                      element.getAttribute('personId') || "",
                      element.getAttribute('parentId') || "",
                      element.getAttribute('done') === '1',
                      element.textContent || ""
                    )
          this.comments[id].partName = partName
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
      const personList = doc.getElementsByTagName('person')
      for (let i = 0; i < personList.length; i++) {
        const element = personList[i]
        const id = element.getAttribute('id') || ""
        this.persons[id] = new Person(
                      element.getAttribute('displayName') || "",
                      element.getAttribute('id') || "",
                      element.getAttribute('providerId') || ""
                  )
      }
    } catch (e) {
      console.log(e)
    }

    return true
  }

  dumpComments() {
    for (let sid in this.sheets) {
      let sheet = this.sheets[sid]
      if (sheet.threadedComments.length === 0) {
        continue
      }

      console.log(`Comment from sheet ${sheet.name}`)
      sheet.threadedComments.forEach(c => {
        this.dumpComment(1, c)
      })
    }
  }
}