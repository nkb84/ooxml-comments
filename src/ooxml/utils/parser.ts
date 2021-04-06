const fs = require('fs')
const { DOMParser } = require('xmldom')

export class Parser {
  private static instance?: Parser = undefined
  parser: DOMParser

  protected constructor () {
    this.parser = new DOMParser()
  }

  static getInstance (): Parser {
    if (Parser.instance === undefined) {
      Parser.instance = new Parser()
    }

    return Parser.instance!
  }

  public parseFromFile (path: string, mime = 'utf-8'): Document {
    const data = fs.readFileSync(path, mime)
    return this.parser.parseFromString(data, "application/xml")
  }

  public parseFromString (data: string) : Document {
    return this.parser.parseFromString(data, "application/xml")
  }
}

