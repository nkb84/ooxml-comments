import { ThreadedComment } from "../entity/comment";
import { Container } from "../entity/container";

export interface CommentList {
  [key: string]: ThreadedComment
}

export interface ContainerList {
  [key: string]: Container
}

export interface Extractor {
  getCommentList(): Promise<ThreadedComment[]>
  dump():void
}

interface ExtractorConstructor {
  new (path: string): Extractor
}

declare var Service: ExtractorConstructor;