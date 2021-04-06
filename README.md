# OOXML Comment Extractor

This utility extract the comments from docx, xlsx and pptx

## Compile

```shell
npm install
npm run build
# Or
npm run watch
```

## Run the tool

```shell
node output/index.js <docx/xlsx/pptx file path>
```

> Node: the command line is just an example of using this tool

## Using in your app

  ```js
  let service = OOXmlExtractor.from(filePath)
  service.getCommentList().then(commentList => {
    console.log(JSON.stringify(commentList))

    // Do something with comment list
    commentList.forEach(comment => {
      // Do something
    });
  })
  ```
Each comment object contain:
  ```json
  {
    "id": "string: Comment Id",
    "ref": "string: Reference number, not use now",
    "time": "Date: Date when this comment is made",
    "userId": "string: User id who made this comment",
    "parentId": "string: Parent comment Id",
    "done": "boolean: Is this comment resolved?",
    "comment": "string: detail comment",
    "partName": "string: the part in OOXML compressed file which contain this comment",
    "location": "string: Place where user put this comment",
    "children": "Comment[]: for all children comments"
  }
  ```