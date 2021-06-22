import { ThreadedComment } from "../entity/comment";
import { ContentType, } from "../entity/content_type";
import Person from "../entity/person";
import { Relation } from "../entity/relation";
import { Container } from "../entity/container";
import { Parser } from "../utils/parser";
import { CommentList, ContainerList } from "./extractor";
import { BaseExtractor } from "./base_extractor";

const path = require('path')

export class PptxExtractor extends BaseExtractor {
  private presentation?: string // workbook name
  private persons: {[key: string]: Person} = {}
  private sheets: ContainerList = {}  // map from sheet name to sheet object
  private comments: CommentList = {}
  private commentPartNameMap: {[key: string]: string} = {}  // map from comment partname to sheet name

  private rootRelationLoaded: boolean = false
  private presentationRelationLoaded: boolean = false

  private static ONS = "http://schemas.openxmlformats.org/presentationml/2006/main"

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
    return this.presentationRelationLoaded
  }

  loadContent(partName: string, data: string) : boolean {
    const contentType = this.getContentType(partName)
    if (contentType === undefined) {
      throw new Error(`Part ${partName} cannot be identified the content type`)
    }

    switch (contentType) {
      case ContentType.SLIDE_COMMENT_AUTHORS: return this.loadPersons(data); break;
      case ContentType.SLIDE_COMMENTS: return this.loadComments(partName, data); break;
      case ContentType.RELATION: return this.loadRelations(partName, data); break;
      default:
        break;
    }

    return true
  }

  private loadRelations(partName: string, content: string): boolean {
    const wbRelsPath = Relation.getRelsPath(this.presentation)
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
            this.presentation = (relations[i].getAttribute('Target') || "").replace('\\', '/')
          }
        }
        this.rootRelationLoaded = true
      } else if (partName === wbRelsPath && !this.presentationRelationLoaded) {
        // Workbook relations, get all sheets
        const relations = doc.getElementsByTagName('Relationship')
        for (let i = 0; i < relations.length; i++) {
          if (relations[i].getAttribute('Type') === Relation.SLIDE) {
            const starget = (relations[i].getAttribute('Target') || "").replace('\\', '/')
            // Get sheet name
            const sname = starget.split('/').pop()?.split('.').slice(0, -1).join('.')
            // Workbook folder
            const paths = this.presentation!.split('/')
            paths.pop()
            const spartName = path.posix.normalize(paths.join('/') + '/' + starget)
            if (sname === undefined || spartName === undefined) {
              throw new Error(`Sheet name or target name is underfine for ${starget}`)
            }
            this.sheets[sname] = new Container(sname, spartName)
          }
        }

        this.presentationRelationLoaded = true
      } else if (partName in sheetReslPaths && this.rootRelationLoaded && this.presentationRelationLoaded) {
        // sheet
        const relations = doc.getElementsByTagName('Relationship')
        for (let i = 0; i < relations.length; i++) {
          if (relations[i].getAttribute('Type') === Relation.COMMENT) {
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
      c.location = sname + ': ' + c.location
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

    try {
      const doc = parser.parseFromString(content)
      const commentList = doc.getElementsByTagNameNS(PptxExtractor.ONS, 'cm')
      for (let i = 0; i < commentList.length; i++) {
        const element = commentList[i]

        // Get comment position
        const poses = element.getElementsByTagNameNS(PptxExtractor.ONS, 'pos')
        if (poses.length <= 0) {
          throw new Error(`Comment in file ${partName} has no pos element`)
        }
        const positionTxt = {x: poses[0].getAttribute('x'), y: poses[0].getAttribute('y')}

        // Get parent
        const parents = element.getElementsByTagName('p15:parentCm')
        const parentId = (parents.length > 0) ? parents[0].getAttribute('idx') || "" : ""
        const id = element.getAttribute('idx') || ""
        this.comments[id] = new ThreadedComment(
                      element.getAttribute('idx') || "",
                      "", // ref
                      element.getAttribute('dt') || "",
                      element.getAttribute('authorId') || "",
                      parentId, // parent id
                      false,
                      element.textContent || ""
                    )
          this.comments[id].partName = partName
          this.comments[id].location = JSON.stringify(positionTxt)
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
      const personList = doc.getElementsByTagNameNS(PptxExtractor.ONS, 'cmAuthor')
      for (let i = 0; i < personList.length; i++) {
        const element = personList[i]
        const id = element.getAttribute('id') || ""
        this.persons[id] = new Person(
                      element.getAttribute('name') || "",
                      element.getAttribute('id') || "",
                      ""
                  )
      }
    } catch (e) {
      console.log(e)
    }

    return true
  }

  private getUserDisplayName (userId: string): String {
    return this.persons.hasOwnProperty(userId) ? this.persons[userId].displayName : userId
  }

  private dumpPersons() {
    for (let pid in this.persons) {
      const p = this.persons[pid]
      console.log(`- User ${p.id} = ${p.displayName}`)
    }
  }

  private dumpComment(idx: number, comment: ThreadedComment) {
    console.log(`${' '.repeat(idx*2)}- Comment ${comment.id}: ${comment.ref} at ${comment.time} by ${this.getUserDisplayName(comment.userId)}`)
    if (comment.children) {
      console.log(`${' '.repeat(idx*2 + 2)}| ${JSON.stringify(comment.comment)}`)
      comment.children.forEach(c => {
        this.dumpComment(idx + 1, c)
      })
    } else {
      console.log(`${' '.repeat(idx*2 + 2)}  ${JSON.stringify(comment.comment)}`)
    }
  }

  private dumpComments() {
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

  public dump(): void {
    this.dumpPersons()
    this.dumpComments()
  }
}