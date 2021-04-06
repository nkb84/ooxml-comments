import { Entry } from "unzipper";
import { ThreadedComment } from "../entity/comment";
import { ContentType, ContentTypes } from "../entity/content_type";
import Person from "../entity/person";
import { Relation } from "../entity/relation";
import { Container } from "../entity/container";
import { Parser } from "../utils/parser";
import { CommentList, ContainerList, Extractor } from "./extractor";

const path = require('path')
const fs = require('fs')
const unziper = require('unzipper')

export class XlsxExtractor implements Extractor {
  workbook?: string // workbook name
  contentTypes: ContentTypes = new ContentTypes()
  persons: {[key: string]: Person} = {}
  sheets: ContainerList = {}  // map from sheet name to sheet object
  comments: CommentList = {}
  commentPartNameMap: {[key: string]: string} = {}  // map from comment partname to sheet name

  private contentTypesLoaded: boolean = false
  private rootRelationLoaded: boolean = false
  private workbookRelationLoaded: boolean = false

  private filePath: string
  constructor (path: string) {
    this.filePath = path
  }

  public getCommentList(): Promise<ThreadedComment[]> {
    let pendingList: {[key: string]: string} = {}

    return new Promise<ThreadedComment[]>(resolve => {
      fs.createReadStream(this.filePath)
      .pipe(unziper.Parse())
      .on('entry', async (entry: Entry) => {
        const data = await entry.buffer()
        if (entry.path === '[Content_Types].xml') {
          // Load content type
          this.loadContentType(data.toString())
          // console.log(`Finish loading content types`)
          // xlsxService.dump()
        } else {
          const contentType = this.getContentType(entry.path)
          if (contentType === undefined) {
            if (this.isContentTypesLoaded()) {
              throw new Error(`Entry ${entry.path} could not identify the content type`)
            }

            // Maybe file '[Content_Types].xml' has not been loaded, add to pending list
            pendingList[entry.path] = data.toString()
            // console.log(`Adding ${entry.path} to pending list`)
          } else {
            // console.log(`Loading ${entry.path} with: ${contentType}`)
            if (!this.loadContent(entry.path, data.toString())) {
              pendingList[entry.path] = data.toString()
              // console.log(`....Adding ${entry.path} to pending list`)
            }
          }
        }
        entry.autodrain()
      })
      .on('finish', () => {
        // console.log(`Pending list has ${JSON.stringify(Object.keys(pendingList))}`)
        // Load pending data files
        let retry = 0
        while (retry < 10 && Object.keys(pendingList).length > 0) {
          for (let partName in pendingList) {
            // console.log(`Loading ${partName} from pending list`)
            if (this.loadContent(partName, pendingList[partName])) {
              // Remove from pending list
              delete pendingList[partName]
            }
          }
          retry ++
        }

        // Dispose pending list
        pendingList = {}

        // Update comment list
        this.updateThreadComments()

        resolve(this.adjustCommentList())
      })
    })
  }

  private adjustCommentList(): ThreadedComment[] {
    const result: ThreadedComment[] = []
    Object.values(this.sheets).forEach(s => {
      if (s.threadedComments.length > 0) {
        result.push(...s.threadedComments)
      }
    })
    return result
  }

  private loadContentType(content: string) {
    const parser = Parser.getInstance()

    try {
      const doc = parser.parseFromString(content)
      // Default
      const defaults = doc.getElementsByTagName('Default')
      for (let i = 0; i < defaults.length; i++) {
        this.contentTypes.addDefault(
                      defaults[i].getAttribute('Extension') || "",
                      defaults[i].getAttribute('ContentType') || ""
                    )
      }

      // Override
      const overrides = doc.getElementsByTagName('Override')
      for (let i = 0; i < overrides.length; i++) {
        this.contentTypes.addOverride(
                      overrides[i].getAttribute('PartName') || "",
                      overrides[i].getAttribute('ContentType') || ""
                    )
      }
    } catch (e) {
      console.log(e)
    }

    this.contentTypesLoaded = true
  }

  private isContentTypesLoaded(): boolean {
    return this.contentTypesLoaded
  }

  private isRootRelationLoaded(): boolean {
    return this.rootRelationLoaded
  }

  private isWorkbookRelationLoaded(): boolean {
    return this.workbookRelationLoaded
  }

  private getContentType(partName: string): ContentType | undefined {
    return this.contentTypes.getContentType(partName)
  }

  private loadContent(partName: string, data: string) : boolean {
    const contentType = this.getContentType(partName)
    if (contentType === undefined) {
      throw new Error(`Part ${partName} cannot be identified the content type`)
    }

    switch (contentType) {
      case ContentType.EXCEL_PERSON: return this.loadPersons(data); break;
      case ContentType.EXCEL_THREADED_COMMENT: return this.loadComments(partName, data); break;
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
          if (relations[i].getAttribute('Type') === Relation.THREADED_COMMENT) {
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

  private updateThreadComments() {
    for (let id in this.comments) {
      const c = this.comments[id]
      const sname = this.getSheetNameFromCommentPart(c.partName)
      if (sname === undefined) {
        throw new Error(`Sheet is empty in comment ${c.id}`)
      }

      // Update location for comment
      const sheet = this.sheets[sname]
      c.location = sname
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

  private getUserById (userId: string): Person | undefined {
    return this.persons[userId]
  }

  private dumpContentTypes () {
    this.contentTypes.defaults.forEach(d => {
      console.log(`- Extension ${d.extension} = ${d.contentType}`)
    })

    this.contentTypes.overrides.forEach(d => {
      console.log(`- Override ${d.partName} = ${d.contentType}`)
    })
  }

  private dumpPersons() {
    for (let pid in this.persons) {
      const p = this.persons[pid]
      console.log(`- User ${p.id} = ${p.displayName}`)
    }
  }

  private dumpComment(idx: number, comment: ThreadedComment) {
    console.log(`${' '.repeat(idx*2)}- Comment ${comment.id}: ${comment.ref} at ${comment.time} by ${this.getUserById(comment.userId)?.displayName}`)
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

  public dump():void {
    this.dumpPersons()
    this.dumpComments()
  }
}