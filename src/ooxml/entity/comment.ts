export class ThreadedComment {
  ref: string
  time: string
  id: string
  userId: string
  done: boolean
  parentId: string
  comment: string
  children?: ThreadedComment[]
  partName?: string
  location?: string

  constructor (id: string, ref: string, time: string, userId: string, parentId: string, done: boolean, comment: string) {
    this.id = id
    this.ref = ref
    const d = new Date(time)
    this.time = d.toISOString() // [d.getFullYear(), d.getMonth() + 1, d.getDate()].join('-')
    this.userId = userId
    this.parentId = parentId
    this.done = done
    this.comment = comment
  }
}