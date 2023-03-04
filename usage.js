const application = require('/index.js')

WScript.Sleep(5000)
//console.log(() => application.getTitle)
console.log(() => application.getTitle())
console.log(() => application.getWindow())
application.pos({ x: 100, y: 200 }, 300)
application.click()
