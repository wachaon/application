
const { resolve } = require('pathname')
const { existsFileSync } = require('filesystem')
const { compile } = require('csharpscript')
const {
    setData,
    getData
} = require('clipboard')
const {
    move,
    get,
    hwnd,
    windowLeft,
    windowTop,
    windowWidth,
    windowHeight,
    max,
    min,
    normal,
    activateTitle,
    activateHandle,
    title
} = require('window')
const {
    pos,
    click,
    leftDown,
    leftUp,
    rightClick,
    rightDown,
    rightUp,
    scroll
} = require('mouse')
const {
    send,
    press,
    release,
    VK_LBUTTON, // left mouse button
    VK_RBUTTON, // right mouse button
    VK_CANCEL, // Ctrl+Break processing
    VK_MBUTTON, // Middle button on 3-button mouse
    VK_XBUTTON1, // mouse X1 button
    VK_XBUTTON2, // mouse X2 button
    VK_BACK, // Backspace key
    VK_TAB, // Tab key
    VK_CLEAR, // clear key
    VK_RETURN, // Enter key
    VK_SHIFT, // shift key
    VK_CONTROL, // Ctrl key
    VK_MENU, // Alt key
    VK_PAUSE, // Pause key
    VK_CAPITAL, // Caps Lock key
    VK_KANA, // IME kana mode
    VK_HANGEUL, // IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
    VK_HANGUL, // IME hangul mode
    VK_JUNJA, // IME Junja mode
    VK_FINAL, // IME final mode
    VK_HANJA, // IME Hanja mode
    VK_KANJI, // IME kanji mode
    VK_ESCAPE, // Esc key
    VK_CONVERT, // IME conversion
    VK_NONCONVERT, // No IME conversion
    VK_ACCEPT, // IME accept
    VK_MODECHANGE, // IME mode change request
    VK_SPACE, // space key
    VK_PRIOR, // Page Up key
    VK_NEXT, // Page Down key
    VK_END, // End key
    VK_HOME, // home key
    VK_LEFT, // cursor key left
    VK_UP, // cursor key up
    VK_RIGHT, // cursor key right
    VK_DOWN, // cursor key down
    VK_SELECT, // Select key
    VK_PRINT, // Print key
    VK_EXECUTE, // Execute key
    VK_SNAPSHOT, // Print Screen key
    VK_INSERT, // Insert key
    VK_DELETE, // Delete key
    VK_HELP, // Help key
    VK_0, // 0 key
    VK_1, // 1 key
    VK_2, // 2 keys
    VK_3, // 3 key
    VK_4, // 4 key
    VK_5, // 5 key
    VK_6, // 6 key
    VK_7, // 7 key
    VK_8, // 8 key
    VK_9, // 9 key
    VK_A, // A key
    VK_B, // B key
    VK_C, // C key
    VK_D, // D key
    VK_E, // E key
    VK_F, // F key
    VK_G, // G key
    VK_H, // H key
    VK_I, // I key
    VK_J, // J key
    VK_K, // K key
    VK_L, // L key
    VK_M, // M key
    VK_N, // N key
    VK_O, // O key
    VK_P, // P key
    VK_Q, // Q key
    VK_R, // R key
    VK_S, // S key
    VK_T, // T key
    VK_U, // U key
    VK_V, // V key
    VK_W, // W key
    VK_X, // X key
    VK_Y, // Y key
    VK_Z, // Z key
    VK_LWIN, // Left Windows key
    VK_RWIN, // Right Windows key
    VK_APPS, // application key
    VK_SLEEP, // sleep key
    VK_NUMPAD0, // numeric keypad 0
    VK_NUMPAD1, // numeric keypad 1
    VK_NUMPAD2, // numeric keypad 2
    VK_NUMPAD3, // numeric keypad 3
    VK_NUMPAD4, // numeric keypad 4
    VK_NUMPAD5, // numeric keypad 5
    VK_NUMPAD6, // numeric keypad 6
    VK_NUMPAD7, // numeric keypad 7
    VK_NUMPAD8, // numeric keypad 8
    VK_NUMPAD9, // numeric keypad 9
    VK_MULTIPLY, // * key
    VK_ADD, // + key
    VK_SEPARATOR, // Separator key
    VK_SUBTRACT, // -key
    VK_DECIMAL, // . key
    VK_DIVIDE, // / key
    VK_F1, // F1 key
    VK_F2, // F2 key
    VK_F3, // F3 key
    VK_F4, // F4 key
    VK_F5, // F5 key
    VK_F6, // F6 key
    VK_F7, // F7 key
    VK_F8, // F8 key
    VK_F9, // F9 key
    VK_F10, // F10 key
    VK_F11, // F11 key
    VK_F12, // F12 key
    VK_F13, // F13 key
    VK_F14, // F14 key
    VK_F15, // F15 key
    VK_F16, // F16 key
    VK_F17, // F17 key
    VK_F18, // F18 key
    VK_F19, // F19 key
    VK_F20, // F20 key
    VK_F21, // F21 key
    VK_F22, // F22 key
    VK_F23, // F23 key
    VK_F24, // F24 key
    VK_NUMLOCK, // NumLock key
    VK_SCROLL, // ScrollLock key
    VK_LSHIFT, // left shift key
    VK_RSHIFT, // right shift key
    VK_LCONTROL, // Left Ctrl key
    VK_RCONTROL, // Right Ctrl key
    VK_LMENU, // Left Alt key
    VK_RMENU, // Right Alt key
    VK_BROWSER_BACK, // browser back key
    VK_BROWSER_FORWARD, // browser forward key
    VK_BROWSER_REFRESH, // browser refresh key
    VK_BROWSER_STOP, // browser stop key
    VK_BROWSER_SEARCH, // browser search key
    VK_BROWSER_FAVORITES, // browser favorite keys
    VK_BROWSER_HOME, // Browser Home key
    VK_VOLUME_MUTE, // volume mute key
    VK_VOLUME_DOWN, // volume down key
    VK_VOLUME_UP, // volume up key
    VK_MEDIA_NEXT_TRACK, // media next track key
    VK_MEDIA_PREV_TRACK, // media pre-track key
    VK_MEDIA_STOP, // media stop key
    VK_MEDIA_PLAY_PAUSE, // media play/pause key
    VK_LAUNCH_MAIL, // Mail launch key
    VK_LAUNCH_MEDIA_SELECT, // media selection key
    VK_LAUNCH_APP1, // launch key 1
    VK_LAUNCH_APP2, // launch key 2
    VK_ICO_HELP, // ?
    VK_ICO_00, // ?
    VK_PROCESSKEY, // IME PROCESS key
    VK_ICO_CLEAR, // ?
    VK_PACKET, // See MSDN for details
    VK_OEM_RESET, // OEM defined key
    VK_OEM_JUMP, // OEM defined key
    VK_OEM_PA1, // OEM defined key
    VK_OEM_PA2, // OEM defined key
    VK_OEM_PA3, // OEM defined key
    VK_OEM_WSCTRL, // OEM defined key
    VK_OEM_CUSEL, // OEM defined key
    VK_OEM_ATTN, // OEM defined key
    VK_OEM_FINISH, // OEM defined key
    VK_OEM_COPY, // OEM defined key
    VK_OEM_AUTO, // OEM defined key
    VK_OEM_ENLW, // OEM defined key
    VK_OEM_BACKTAB, // OEM defined key
    VK_ATTN, // Attn key
    VK_CRSEL, // CrSel key
    VK_EXSEL, // Exsel key
    VK_EREOF, // Erase EOF key
    VK_PLAY, // Play key
    VK_ZOOM, // Zoom key
    VK_NONAME, // Reserved
    VK_PA1, // PA1 key
    VK_OEM_CLEAR // clear key
} = require('keyboard')

const clipboard = [resolve(__dirname, 'wes_modules/clipboard/clipboard.exe'), resolve(__dirname, 'wes_modules/clipboard/src/clipboard.cs')]
const window = [resolve(__dirname, 'wes_modules/window/window.exe'), resolve(__dirname, 'wes_modules/window/src/window.cs')]
const mouse = [resolve(__dirname, 'wes_modules/mouse/mouse.exe'), resolve(__dirname, 'wes_modules/mouse/src/mouse.cs')]
const keyboard = [resolve(__dirname, 'wes_modules/keyboard/keyboard.exe'), resolve(__dirname, 'wes_modules/keyboard/src/keyboard.cs')]

const application = [clipboard, window, mouse, keyboard]
application.forEach(app => {
    if (!existsFileSync(app[0])) compile(app[1], { out: app[0] })
})

module.exports = {
    "setData": function application_setData(text = '', time = 0) {
        setData(text)
        WScript.Sleep(time)
    },
    "getData": function application_getData(time = 0) {
        const res = getData()
        WScript.Sleep(time)
        return res
    },
    "move": function application_move(position, time = 0) {
        const { left, top, width, height } = position
        move(left, top, width, height)
        WScript.Sleep(time)
    },
    "getWindow": function application_get(time = 0) {
        const exp = /(\d+)/
        const res = {
            hwnd: parseInt(hwnd(), 10),
            title: title(),
            left: parseInt(windowLeft().match(exp)[1], 10),
            top: parseInt(windowTop().match(exp)[1], 10),
            width: parseInt(windowWidth().match(exp)[1], 10),
            height: parseInt(windowHeight().match(exp)[1], 10)

        }
        WScript.Sleep(time)
        return res
    },
    "getHwnd": function application_hwnd(time = 0) {
        const res = hwnd()
        WScript.Sleep(time)
        return res
    },
    "getLeft": function application_windowLeft(time = 0) {
        const res = windowLeft()
        WScript.Sleep(time)
        return res
    },
    "getTop": function application_windowTop(time = 0) {
        const res = windowTop()
        WScript.Sleep(time)
        return res
    },
    "getWidth": function application_windowWidth(time = 0) {
        const res = windowWidth()
        WScript.Sleep(time)
        return res
    },
    "getHeight": function application_windowHeight(time = 0) {
        const res = windowHeight()
        WScript.Sleep(time)
        return res
    },
    "max": function application_max(time = 0) {
        max()
        WScript.Sleep(time)
    },
    "min": function application_min(time = 0) {
        min()
        WScript.Sleep(time)
    },
    "normal": function application_normal(time = 0) {
        normal()
        WScript.Sleep(time)
    },
    "activateTitle": function application_activateTitle(title, time = 0) {
        activateTitle(title)
        WScript.Sleep(time)
    },
    "activateHandle": function application_activateHandle(hWnd, time = 0) {
        activateHandle(hWnd)
        WScript.Sleep(time)
    },
    "getTitle": function application_title(time = 0) {
        const res = title()
        WScript.Sleep(time)
        return res
    },
    "pos": function application_pos(position, time = 0) {
        const { x, y } = position
        pos(x, y)
        WScript.Sleep(time)
    },
    "click": function application_click(time = 0) {
        click()
        WScript.Sleep(time)
    },
    "leftDown": function application_leftDown(time = 0) {
        leftDown()
        WScript.Sleep(time)
    },
    "leftUp": function application_leftUp(time = 0) {
        leftUp()
        WScript.Sleep(time)
    },
    "rightClick": function application_rightClick(time = 0) {
        rightClick()
        WScript.Sleep(time)
    },
    "rightDown": function application_rightDown(time = 0) {
        rightDown()
        WScript.Sleep(time)
    },
    "rightUp": function application_rightUp(time = 0) {
        rightUp()
        WScript.Sleep(time)
    },
    "scroll": function application_scroll(movement = 0, time = 0) {
        scroll(movement)
        WScript.Sleep(time)
    },
    "send": function application_send(keyCode, time = 0) {
        send(keyCode)
        WScript.Sleep(time)
    },
    "press": function application_press(keyCode, time = 0) {
        press(keyCode)
        WScript.Sleep(time)
    },
    "release": function application_release(keyCode, time = 0) {
        release(keyCode)
        WScript.Sleep(time)
    },
    VK_LBUTTON, // left mouse button
    VK_RBUTTON, // right mouse button
    VK_CANCEL, // Ctrl+Break processing
    VK_MBUTTON, // Middle button on 3-button mouse
    VK_XBUTTON1, // mouse X1 button
    VK_XBUTTON2, // mouse X2 button
    VK_BACK, // Backspace key
    VK_TAB, // Tab key
    VK_CLEAR, // clear key
    VK_RETURN, // Enter key
    VK_SHIFT, // shift key
    VK_CONTROL, // Ctrl key
    VK_MENU, // Alt key
    VK_PAUSE, // Pause key
    VK_CAPITAL, // Caps Lock key
    VK_KANA, // IME kana mode
    VK_HANGEUL, // IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
    VK_HANGUL, // IME hangul mode
    VK_JUNJA, // IME Junja mode
    VK_FINAL, // IME final mode
    VK_HANJA, // IME Hanja mode
    VK_KANJI, // IME kanji mode
    VK_ESCAPE, // Esc key
    VK_CONVERT, // IME conversion
    VK_NONCONVERT, // No IME conversion
    VK_ACCEPT, // IME accept
    VK_MODECHANGE, // IME mode change request
    VK_SPACE, // space key
    VK_PRIOR, // Page Up key
    VK_NEXT, // Page Down key
    VK_END, // End key
    VK_HOME, // home key
    VK_LEFT, // cursor key left
    VK_UP, // cursor key up
    VK_RIGHT, // cursor key right
    VK_DOWN, // cursor key down
    VK_SELECT, // Select key
    VK_PRINT, // Print key
    VK_EXECUTE, // Execute key
    VK_SNAPSHOT, // Print Screen key
    VK_INSERT, // Insert key
    VK_DELETE, // Delete key
    VK_HELP, // Help key
    VK_0, // 0 key
    VK_1, // 1 key
    VK_2, // 2 keys
    VK_3, // 3 key
    VK_4, // 4 key
    VK_5, // 5 key
    VK_6, // 6 key
    VK_7, // 7 key
    VK_8, // 8 key
    VK_9, // 9 key
    VK_A, // A key
    VK_B, // B key
    VK_C, // C key
    VK_D, // D key
    VK_E, // E key
    VK_F, // F key
    VK_G, // G key
    VK_H, // H key
    VK_I, // I key
    VK_J, // J key
    VK_K, // K key
    VK_L, // L key
    VK_M, // M key
    VK_N, // N key
    VK_O, // O key
    VK_P, // P key
    VK_Q, // Q key
    VK_R, // R key
    VK_S, // S key
    VK_T, // T key
    VK_U, // U key
    VK_V, // V key
    VK_W, // W key
    VK_X, // X key
    VK_Y, // Y key
    VK_Z, // Z key
    VK_LWIN, // Left Windows key
    VK_RWIN, // Right Windows key
    VK_APPS, // application key
    VK_SLEEP, // sleep key
    VK_NUMPAD0, // numeric keypad 0
    VK_NUMPAD1, // numeric keypad 1
    VK_NUMPAD2, // numeric keypad 2
    VK_NUMPAD3, // numeric keypad 3
    VK_NUMPAD4, // numeric keypad 4
    VK_NUMPAD5, // numeric keypad 5
    VK_NUMPAD6, // numeric keypad 6
    VK_NUMPAD7, // numeric keypad 7
    VK_NUMPAD8, // numeric keypad 8
    VK_NUMPAD9, // numeric keypad 9
    VK_MULTIPLY, // * key
    VK_ADD, // + key
    VK_SEPARATOR, // Separator key
    VK_SUBTRACT, // -key
    VK_DECIMAL, // . key
    VK_DIVIDE, // / key
    VK_F1, // F1 key
    VK_F2, // F2 key
    VK_F3, // F3 key
    VK_F4, // F4 key
    VK_F5, // F5 key
    VK_F6, // F6 key
    VK_F7, // F7 key
    VK_F8, // F8 key
    VK_F9, // F9 key
    VK_F10, // F10 key
    VK_F11, // F11 key
    VK_F12, // F12 key
    VK_F13, // F13 key
    VK_F14, // F14 key
    VK_F15, // F15 key
    VK_F16, // F16 key
    VK_F17, // F17 key
    VK_F18, // F18 key
    VK_F19, // F19 key
    VK_F20, // F20 key
    VK_F21, // F21 key
    VK_F22, // F22 key
    VK_F23, // F23 key
    VK_F24, // F24 key
    VK_NUMLOCK, // NumLock key
    VK_SCROLL, // ScrollLock key
    VK_LSHIFT, // left shift key
    VK_RSHIFT, // right shift key
    VK_LCONTROL, // Left Ctrl key
    VK_RCONTROL, // Right Ctrl key
    VK_LMENU, // Left Alt key
    VK_RMENU, // Right Alt key
    VK_BROWSER_BACK, // browser back key
    VK_BROWSER_FORWARD, // browser forward key
    VK_BROWSER_REFRESH, // browser refresh key
    VK_BROWSER_STOP, // browser stop key
    VK_BROWSER_SEARCH, // browser search key
    VK_BROWSER_FAVORITES, // browser favorite keys
    VK_BROWSER_HOME, // Browser Home key
    VK_VOLUME_MUTE, // volume mute key
    VK_VOLUME_DOWN, // volume down key
    VK_VOLUME_UP, // volume up key
    VK_MEDIA_NEXT_TRACK, // media next track key
    VK_MEDIA_PREV_TRACK, // media pre-track key
    VK_MEDIA_STOP, // media stop key
    VK_MEDIA_PLAY_PAUSE, // media play/pause key
    VK_LAUNCH_MAIL, // Mail launch key
    VK_LAUNCH_MEDIA_SELECT, // media selection key
    VK_LAUNCH_APP1, // launch key 1
    VK_LAUNCH_APP2, // launch key 2
    VK_ICO_HELP, // ?
    VK_ICO_00, // ?
    VK_PROCESSKEY, // IME PROCESS key
    VK_ICO_CLEAR, // ?
    VK_PACKET, // See MSDN for details
    VK_OEM_RESET, // OEM defined key
    VK_OEM_JUMP, // OEM defined key
    VK_OEM_PA1, // OEM defined key
    VK_OEM_PA2, // OEM defined key
    VK_OEM_PA3, // OEM defined key
    VK_OEM_WSCTRL, // OEM defined key
    VK_OEM_CUSEL, // OEM defined key
    VK_OEM_ATTN, // OEM defined key
    VK_OEM_FINISH, // OEM defined key
    VK_OEM_COPY, // OEM defined key
    VK_OEM_AUTO, // OEM defined key
    VK_OEM_ENLW, // OEM defined key
    VK_OEM_BACKTAB, // OEM defined key
    VK_ATTN, // Attn key
    VK_CRSEL, // CrSel key
    VK_EXSEL, // Exsel key
    VK_EREOF, // Erase EOF key
    VK_PLAY, // Play key
    VK_ZOOM, // Zoom key
    VK_NONAME, // Reserved
    VK_PA1, // PA1 key
    VK_OEM_CLEAR // clear key
}
