export class Relation {
  public static EXTENDED_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
  public static CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
  public static THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
  public static OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  public static STYLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  public static THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
  public static WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  public static PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person"
  public static PEOPLE = "http://schemas.microsoft.com/office/2011/relationships/people"
  public static SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
  public static THREADED_COMMENT = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
  public static COMMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  public static MS_COMMENT = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
  public static COMMENT_EXTENDED = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
  public static COMMENT_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
  public static COMMENT_EXTENSIBLE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
  public static COMMENT_AUTHORS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
  public static VMLDRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"

  public static SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
  public static SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
  public static VIEW_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
  public static PRES_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
  public static HANDOUT_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"
  public static TABLE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
  public static NODES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"

  id: string
  type: string
  target: string

  constructor(id: string, type: string, target: string) {
    this.id = id
    this.type = type
    this.target = target
  }

  static getRelsPath(target?: string): string | undefined {
    if (target === undefined) {
      return undefined
    }

    const paths = target.split('/')
    const wname = paths.pop()

    return [...paths,
            '_rels',
            wname + '.rels'].join('/')
  }
}