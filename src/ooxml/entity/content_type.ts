export enum ContentType {
  RELATION = "application/vnd.openxmlformats-package.relationships+xml",
  DRAWING = "application/vnd.openxmlformats-officedocument.vmlDrawing",
  XML = "application/xml",
  CORE_PROPERTIES = "application/vnd.openxmlformats-package.core-properties+xml",
  EXTENDED_PROPERTIES = "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  THEME = "application/vnd.openxmlformats-officedocument.theme+xml",

  IMG_EMF = "image/x-emf",
  IMG_JPEG = "image/jpeg",
  IMG_PNG = "image/png",

  WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  EXCEL_COMMENT = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
  EXCEL_THREADED_COMMENT = "application/vnd.ms-excel.threadedcomments+xml",
  EXCEL_PERSON = "application/vnd.ms-excel.person+xml",

  DOCUMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  WORD_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  WORD_SETTING = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  WORD_WEBSETTING = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
  WORD_COMMENT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
  WORD_COMMENT_EXTENDED = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
  WORD_COMMENT_IDS = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
  WORD_COMMENT_EXTENSIBLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
  WORD_FONTTABLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
  WORD_PEOPLE = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",

  PRESENTATION = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  SLIDE_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
  SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
  NODES_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
  SLIDE_HANDOUT_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml",
  SLIDE_COMMENT_AUTHORS = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml",
  SLIDE_PRES_PROPERTIES = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml",
  SLIDE_VIEW_PROPERTIES = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
  SLIDE_TABLE_STYLE = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml",
  SLIDE_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
  NODES_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  SLIDE_COMMENTS = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
}

interface Default {
  extension: string
  contentType: string
}

interface Override {
  partName: string
  contentType: string
}

export class ContentTypes {
  defaults: Default[] = []
  overrides: Override[] = []

  public addDefault(extension: string, contentType: string) {
    this.defaults.push({extension, contentType})
  }

  public addOverride(partName: string, contentType: string) {
    this.overrides.push({partName, contentType})
  }

  public getContentType(pathName: string): ContentType | undefined {
    // Search in override
    let idx = this.overrides.findIndex(o => o.partName.includes(pathName))
    if (idx >= 0) {
      return this.overrides[idx].contentType as ContentType
    }

    // Search in defaults. First, get file extension
    const ext = pathName.split('.').pop()?.toLocaleLowerCase()
    if (ext !== undefined) {
      idx = this.defaults.findIndex(d => d.extension === ext)
      if (idx >= 0) {
        return this.defaults[idx].contentType as ContentType
      }
    }

    return undefined
  }
}