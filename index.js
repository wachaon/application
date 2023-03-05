const WShell = require('WScript.Shell')
const { execCommand, forEach } = require('utility')
const { isNumber } = require('typecheck')
const { resolve, toWin32Sep } = require('pathname')
const { existsFileSync } = require('filesystem')
const isCLI = require('isCLI')

const application = resolve(__dirname, 'application.exe')
const source = resolve(__dirname, 'src/application.cs')
if (!existsFileSync(application)) compile(source, { out: application })

module.exports = {
    move, getState, normalize, minimize, maximize, activate, send, press, release, pos, click, leftDown, leftUp,
    rightClick, rightDown, rightUp, middleClick, middleDown, middleUp, scroll, setClipboard, getClipboard,
    VK_LBUTTON: 0x01, // left mouse button
    VK_RBUTTON: 0x02, // right mouse button
    VK_CANCEL: 0x03, // Ctrl+Break processing
    VK_MBUTTON: 0x04, // Middle button on 3-button mouse
    VK_XBUTTON1: 0x05, // mouse X1 button
    VK_XBUTTON2: 0x06, // mouse X2 button
    VK_BACK: 0x08, // Backspace key
    VK_TAB: 0x09, // Tab key
    VK_CLEAR: 0x0C, // clear key
    VK_RETURN: 0x0D, // Enter key
    VK_SHIFT: 0x10, // shift key
    VK_CONTROL: 0x11, // Ctrl key
    VK_MENU: 0x12, // Alt key
    VK_PAUSE: 0x13, // Pause key
    VK_CAPITAL: 0x14, // Caps Lock key
    VK_KANA: 0x15, // IME kana mode
    VK_HANGEUL: 0x15, // IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
    VK_HANGUL: 0x15, // IME hangul mode
    VK_JUNJA: 0x17, // IME Junja mode
    VK_FINAL: 0x18, // IME final mode
    VK_HANJA: 0x19, // IME Hanja mode
    VK_KANJI: 0x19, // IME kanji mode
    VK_ESCAPE: 0x1B, // Esc key
    VK_CONVERT: 0x1C, // IME conversion
    VK_NONCONVERT: 0x1D, // No IME conversion
    VK_ACCEPT: 0x1E, // IME accept
    VK_MODECHANGE: 0x1F, // IME mode change request
    VK_SPACE: 0x20, // space key
    VK_PRIOR: 0x21, // Page Up key
    VK_NEXT: 0x22, // Page Down key
    VK_END: 0x23, // End key
    VK_HOME: 0x24, // home key
    VK_LEFT: 0x25, // cursor key left
    VK_UP: 0x26, // cursor key up
    VK_RIGHT: 0x27, // cursor key right
    VK_DOWN: 0x28, // cursor key down
    VK_SELECT: 0x29, // Select key
    VK_PRINT: 0x2A, // Print key
    VK_EXECUTE: 0x2B, // Execute key
    VK_SNAPSHOT: 0x2C, // Print Screen key
    VK_INSERT: 0x2D, // Insert key
    VK_DELETE: 0x2E, // Delete key
    VK_HELP: 0x2F, // Help key
    VK_0: 0x30, // 0 key
    VK_1: 0x31, // 1 key
    VK_2: 0x32, // 2 keys
    VK_3: 0x33, // 3 key
    VK_4: 0x34, // 4 key
    VK_5: 0x35, // 5 key
    VK_6: 0x36, // 6 key
    VK_7: 0x37, // 7 key
    VK_8: 0x38, // 8 key
    VK_9: 0x39, // 9 key
    VK_A: 0x41, // A key
    VK_B: 0x42, // B key
    VK_C: 0x43, // C key
    VK_D: 0x44, // D key
    VK_E: 0x45, // E key
    VK_F: 0x46, // F key
    VK_G: 0x47, // G key
    VK_H: 0x48, // H key
    VK_I: 0x49, // I key
    VK_J: 0x4A, // J key
    VK_K: 0x4B, // K key
    VK_L: 0x4C, // L key
    VK_M: 0x4D, // M key
    VK_N: 0x4E, // N key
    VK_O: 0x4F, // O key
    VK_P: 0x50, // P key
    VK_Q: 0x51, // Q key
    VK_R: 0x52, // R key
    VK_S: 0x53, // S key
    VK_T: 0x54, // T key
    VK_U: 0x55, // U key
    VK_V: 0x56, // V key
    VK_W: 0x57, // W key
    VK_X: 0x58, // X key
    VK_Y: 0x59, // Y key
    VK_Z: 0x5A, // Z key
    VK_LWIN: 0x5B, // Left Windows key
    VK_RWIN: 0x5C, // Right Windows key
    VK_APPS: 0x5D, // application key
    VK_SLEEP: 0x5F, // sleep key
    VK_NUMPAD0: 0x60, // numeric keypad 0
    VK_NUMPAD1: 0x61, // numeric keypad 1
    VK_NUMPAD2: 0x62, // numeric keypad 2
    VK_NUMPAD3: 0x63, // numeric keypad 3
    VK_NUMPAD4: 0x64, // numeric keypad 4
    VK_NUMPAD5: 0x65, // numeric keypad 5
    VK_NUMPAD6: 0x66, // numeric keypad 6
    VK_NUMPAD7: 0x67, // numeric keypad 7
    VK_NUMPAD8: 0x68, // numeric keypad 8
    VK_NUMPAD9: 0x69, // numeric keypad 9
    VK_MULTIPLY: 0x6A, // * key
    VK_ADD: 0x6B, // + key
    VK_SEPARATOR: 0x6C, // Separator key
    VK_SUBTRACT: 0x6D, // -key
    VK_DECIMAL: 0x6E, // . key
    VK_DIVIDE: 0x6F, // / key
    VK_F1: 0x70, // F1 key
    VK_F2: 0x71, // F2 key
    VK_F3: 0x72, // F3 key
    VK_F4: 0x73, // F4 key
    VK_F5: 0x74, // F5 key
    VK_F6: 0x75, // F6 key
    VK_F7: 0x76, // F7 key
    VK_F8: 0x77, // F8 key
    VK_F9: 0x78, // F9 key
    VK_F10: 0x79, // F10 key
    VK_F11: 0x7A, // F11 key
    VK_F12: 0x7B, // F12 key
    VK_F13: 0x7C, // F13 key
    VK_F14: 0x7D, // F14 key
    VK_F15: 0x7E, // F15 key
    VK_F16: 0x7F, // F16 key
    VK_F17: 0x80, // F17 key
    VK_F18: 0x81, // F18 key
    VK_F19: 0x82, // F19 key
    VK_F20: 0x83, // F20 key
    VK_F21: 0x84, // F21 key
    VK_F22: 0x85, // F22 key
    VK_F23: 0x86, // F23 key
    VK_F24: 0x87, // F24 key
    VK_NUMLOCK: 0x90, // NumLock key
    VK_SCROLL: 0x91, // ScrollLock key
    VK_LSHIFT: 0xA0, // left shift key
    VK_RSHIFT: 0xA1, // right shift key
    VK_LCONTROL: 0xA2, // Left Ctrl key
    VK_RCONTROL: 0xA3, // Right Ctrl key
    VK_LMENU: 0xA4, // Left Alt key
    VK_RMENU: 0xA5, // Right Alt key
    VK_BROWSER_BACK: 0xA6, // browser back key
    VK_BROWSER_FORWARD: 0xA7, // browser forward key
    VK_BROWSER_REFRESH: 0xA8, // browser refresh key
    VK_BROWSER_STOP: 0xA9, // browser stop key
    VK_BROWSER_SEARCH: 0xAA, // browser search key
    VK_BROWSER_FAVORITES: 0xAB, // browser favorite keys
    VK_BROWSER_HOME: 0xAC, // Browser Home key
    VK_VOLUME_MUTE: 0xAD, // volume mute key
    VK_VOLUME_DOWN: 0xAE, // volume down key
    VK_VOLUME_UP: 0xAF, // volume up key
    VK_MEDIA_NEXT_TRACK: 0xB0, // media next track key
    VK_MEDIA_PREV_TRACK: 0xB1, // media pre-track key
    VK_MEDIA_STOP: 0xB2, // media stop key
    VK_MEDIA_PLAY_PAUSE: 0xB3, // media play/pause key
    VK_LAUNCH_MAIL: 0xB4, // Mail launch key
    VK_LAUNCH_MEDIA_SELECT: 0xB5, // media selection key
    VK_LAUNCH_APP1: 0xB6, // launch key 1
    VK_LAUNCH_APP2: 0xB7, // launch key 2
    VK_ICO_HELP: 0xE3, // ?
    VK_ICO_00: 0xE4, // ?
    VK_PROCESSKEY: 0xE5, // IME PROCESS key
    VK_ICO_CLEAR: 0xE6, // ?
    VK_PACKET: 0xE7, // See MSDN for details
    VK_OEM_RESET: 0xE9, // OEM defined key
    VK_OEM_JUMP: 0xEA, // OEM defined key
    VK_OEM_PA1: 0xEB, // OEM defined key
    VK_OEM_PA2: 0xEC, // OEM defined key
    VK_OEM_PA3: 0xED, // OEM defined key
    VK_OEM_WSCTRL: 0xEE, // OEM defined key
    VK_OEM_CUSEL: 0xEF, // OEM defined key
    VK_OEM_ATTN: 0xF0, // OEM defined key
    VK_OEM_FINISH: 0xF1, // OEM defined key
    VK_OEM_COPY: 0xF2, // OEM defined key
    VK_OEM_AUTO: 0xF3, // OEM defined key
    VK_OEM_ENLW: 0xF4, // OEM defined key
    VK_OEM_BACKTAB: 0xF5, // OEM defined key
    VK_ATTN: 0xF6, // Attn key
    VK_CRSEL: 0xF7, // CrSel key
    VK_EXSEL: 0xF8, // Exsel key
    VK_EREOF: 0xF9, // Erase EOF key
    VK_PLAY: 0xFA, // Play key
    VK_ZOOM: 0xFB, // Zoom key
    VK_NONAME: 0xFC, // Reserved
    VK_PA1: 0xFD, // PA1 key
    VK_OEM_CLEAR: 0xFE // clear key
}

function move(rect = {}, time = 0) {
    const left = rect.left || 0
    const top = rect.top || 0
    const width = rect.width || 200
    const height = rect.height || 100

    execCommand(`${application} window.move ${left} ${top} ${width} ${height}`)
    WScript.Sleep(time)
}

function getState() {
    const res = execCommand(`${application} window.get`)
    return JSON.parse(res)
}

function normalize(time = 0) {
    execCommand(`${application} window.normal`)
    WScript.Sleep(time)
}

function minimize(time = 0) {
    execCommand(`${application} window.min`)
    WScript.Sleep(time)
}

function maximize(time = 0) {
    execCommand(`${application} window.max`)
    WScript.Sleep(time)
}

function activate(key, time = 0) {
    if (isNumber(key)) execCommand(`${application} window.activate.window.handle ${key}`)
    else execCommand(`${application} window.activate.window.title "${key}"`)
    WScript.Sleep(time)
}

function send(key, time = 0) {
    if (isNumber(key)) execCommand(`${application} keyboard.send ${key}`)
    else WShell.Sendkeys(key)
    WScript.Sleep(time)
}

function press(key, time = 0) {
    execCommand(`${application} keyboard.press ${key}`)
    WScript.Sleep(time)
}

function release(key, time = 0) {
    execCommand(`${application} keyboard.release ${key}`)
    WScript.Sleep(time)
}

function pos(position = {}, time = 0) {
    const left = position.left || position.x || 0
    const top = position.top || position.y || 0
    execCommand(`${application} mouse.pos ${left} ${top}`)
    WScript.Sleep(time)
}

function click(time = 0) {
    execCommand(`${application} mouse.click`)
    WScript.Sleep(time)
}

function leftDown(time = 0) {
    execCommand(`${application} mouse.leftDown`)
    WScript.Sleep(time)
}

function leftUp(time = 0) {
    execCommand(`${application} mouse.leftUp`)
    WScript.Sleep(time)
}

function rightClick(time = 0) {
    execCommand(`${application} mouse.rightClick`)
    WScript.Sleep(time)
}

function rightDown(time = 0) {
    execCommand(`${application} mouse.rightDown`)
    WScript.Sleep(time)
}

function rightUp(time = 0) {
    execCommand(`${application} mouse.rightUp`)
    WScript.Sleep(time)
}

function middleClick(time = 0) {
    execCommand(`${application} mouse.middleClick`)
    WScript.Sleep(time)
}

function middleDown(time = 0) {
    execCommand(`${application} mouse.middleDown`)
    WScript.Sleep(time)
}

function middleUp(time = 0) {
    execCommand(`${application} mouse.middleUp`)
    WScript.Sleep(time)
}

function scroll(movement = 0) {
    execCommand(`${application} mouse.scroll ${movement}`)
    WScript.Sleep(time)
}

function setClipboard(data, type = 'text', time = 0) {
    if (isNumber(type)) [time, type] = [type, 'text']
    if (type.toLowerCase() === 'text') execCommand(`${application} clipboard.setData Text "${data}"`)
    WScript.Sleep(time)
}

function getClipboard(time = 0) {
    const res = execCommand(`${application} clipboard.getData Text`)
    WScript.Sleep(time)
    return res
}

// util
function getCompiler() {
    return [
        "v4.0.30319",
        "V3.5",
        "V3.0",
        "v2.0.50727",
        "v1.1.4322",
        "v1.0.3705"
    ]
        .map((ver) => {
            return toWin32Sep(
                resolve(
                    WShell.ExpandEnvironmentStrings("%SystemRoot%"),
                    "Microsoft.NET/Framework",
                    ver,
                    "csc.exe"
                )
            )
        })
        .find(exe => existsFileSync(exe))
}

function compile(input, options = {}) {
    options.out = toWin32Sep(
        resolve(
            process.cwd(),
            (
                'out' in options
                    ? options.out
                    : basename(input, extname(input)) + ".exe"
            )
        )
    )
    const opts = Object.assign({
        target: "exe"
    }, options)

    let params = []
    forEach((value, key) => {
        params.push(`/${key}:${value}`)
    })(opts)

    const command = `${getCompiler()} ${params.join(" ")} ${toWin32Sep(resolve(process.cwd(), input))}`
    const result = execCommand(command)

    if (isCLI(__filename)) console.log(result)
    else return result
}