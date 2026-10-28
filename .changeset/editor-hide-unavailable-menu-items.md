---
'@office-kit/pptx-editor': minor
---

The menu bar no longer shows items the editor can never run. Before, it listed every item of the reference desktop app's menus and disabled the ones the editor has no counterpart for, or that a web page cannot do (Print, Share, Export, Macro, SmartArt, the online picture and translation services, the whole Window menu, and so on), so users scanned past dozens of greyed-out entries to find the ones that work. Those items, the submenus left empty and the Window menu are now gone, and separators no longer double up. Items that are unavailable only in the current state (Cut with nothing selected, Symbol outside text, Slide Show without a host that presents) stay, disabled with the reason as their tooltip. The shortcuts of hidden items still do nothing rather than reach the browser, so ⌘P does not print the editor page.
