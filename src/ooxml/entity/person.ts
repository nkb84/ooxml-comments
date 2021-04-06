export default class Person {
  displayName: string
  id: string
  providerId: string

  constructor (name: string, id: string, providerId: string) {
    this.displayName = name
    this.id = id
    this.providerId = providerId
  }
}