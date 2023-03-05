const application = require('/index.js')

console.log(() => application.activate("usage.js - application - Visual Studio Code", 200))
application.minimize(50)
application.maximize()
console.log(() => application.getState())
application.pos({ x: 700, y: 500 }, 30)
application.click()

application.setClipboard("こんにちは世界")
console.log(() => application.activate("chrome", 200))
//application.send('^v', 500)

application.activate("wes", 200)
console.log(application.getClipboard())