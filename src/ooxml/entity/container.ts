import { ThreadedComment } from "./comment";
import { Relation } from "./relation";

export class Container {
  name: string
  partname: string
  threadedComments: ThreadedComment[] = []

  constructor (name: string, partName: string) {
    this.name = name
    this.partname = partName
  }

  public getRelsPath(): string | undefined {
    return Relation.getRelsPath(this.partname)
  }
}