import { Entry } from "unzipper";
import { ThreadedComment } from "../entity/comment";
import { ContentType, ContentTypes } from "../entity/content_type";
import Person from "../entity/person";
import { Parser } from "../utils/index";
import { Extractor } from "./extractor";
const fs = require('fs')
const unziper = require('unzipper')

export abstract class BaseExtractor implements Extractor {
  private contentTypes: ContentTypes = new ContentTypes()
  private contentTypesLoaded: boolean = false
  private filePath: string
  protected persons: {[key: string]: Person} = {}

  constructor(path: string) {
    this.filePath = path
  }

  private onEntry(path: string, data: Buffer, pendingList: { [key: string]: string }) {
    if (path === '[Content_Types].xml') {
      // Load content type
      this.loadContentType(data.toString())
      console.log(`Finish loading content types`)
      // xlsxService.dump()
    } else {
      const contentType = this.getContentType(path)
      if (contentType === undefined) {
        if (this.isContentTypesLoaded()) {
          throw new Error(`Entry ${path} could not identify the content type`)
        }

        // Maybe file '[Content_Types].xml' has not been loaded, add to pending list
        pendingList[path] = data.toString()
        // console.log(`Adding ${entry.path} to pending list`)
      } else {
        // console.log(`Loading ${entry.path} with: ${contentType}`)
        if (!this.loadContent(path, data.toString())) {
          pendingList[path] = data.toString()
          // console.log(`....Adding ${entry.path} to pending list`)
        }
      }
    }
  }

  private onFinish(pendingList: { [key: string]: string }) {
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
      retry++
    }

    // Update comment list
    this.updateThreadComments()

    return this.adjustCommentList()
  }

  protected loadContentType(content: string) {
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

  protected getContentType(partName: string): ContentType | undefined {
    return this.contentTypes.getContentType(partName)
  }

  protected isContentTypesLoaded(): boolean {
    return this.contentTypesLoaded
  }

  public getCommentList(): Promise<ThreadedComment[]> {
    let pendingList: { [key: string]: string } = {}
    const entryPromises: Promise<void>[] = []

    return new Promise<ThreadedComment[]>(resolve => {
      fs.createReadStream(this.filePath)
        .pipe(unziper.Parse())
        .on('entry', (entry: Entry) => {
          const promise = entry.buffer()
            .then(data => {
              this.onEntry(entry.path, data, pendingList)
            })
            .catch((error) => {
              console.log(error)
            })
            .finally(() => {
              entry.autodrain()
            })

          entryPromises.push(promise)
        })
        .on('finish', async () => {
          console.log('Finish')

          // Wait for finish loading all files
          await Promise.all(entryPromises)
          const commentList = this.onFinish(pendingList);

          // Dispose pending list
          pendingList = {}
          resolve(commentList)
        })
    })
  }

  protected getUserDisplayName (userId: string): String {
    return this.persons.hasOwnProperty(userId) ? this.persons[userId].displayName : userId
  }

  protected dumpPersons() {
    for (let pid in this.persons) {
      const p = this.persons[pid]
      console.log(`- User ${p.id} = ${p.displayName}`)
    }
  }

  protected dumpContentTypes() {
    this.contentTypes.defaults.forEach(d => {
      console.log(`- Extension ${d.extension} = ${d.contentType}`)
    })

    this.contentTypes.overrides.forEach(d => {
      console.log(`- Override ${d.partName} = ${d.contentType}`)
    })
  }

  protected dumpComment(idx: number, comment: ThreadedComment) {
    console.log(`${' '.repeat(idx*2)}- Comment ${comment.id}: ${comment.ref} at ${comment.time} by ${this.getUserDisplayName(comment.userId)} ${comment.done ? "done" : ""}`)
    if (comment.children) {
      console.log(`${' '.repeat(idx*2 + 2)}| ${JSON.stringify(comment.comment)}`)
      comment.children.forEach(c => {
        this.dumpComment(idx + 1, c)
      })
    } else {
      console.log(`${' '.repeat(idx*2 + 2)}  ${JSON.stringify(comment.comment)}`)
    }
  }

  dump():void {
    this.dumpPersons()
    this.dumpComments()
  }

  abstract loadContent(partName: string, data: string): boolean
  abstract updateThreadComments(): void
  abstract adjustCommentList(): ThreadedComment[]
  abstract dumpComments():void
}