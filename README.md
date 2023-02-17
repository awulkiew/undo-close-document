# Undo Close Document
### Extension for Visual Studio 2013, 2015, 2017 and 2019

[![VS Marketplace](https://vsmarketplacebadges.dev/version-short/AdamWulkiewicz.UndoCloseDocument.svg)](https://marketplace.visualstudio.com/items?itemName=AdamWulkiewicz.UndoCloseDocument)
[![Installs](https://vsmarketplacebadges.dev/installs-short/AdamWulkiewicz.UndoCloseDocument.svg)](https://marketplace.visualstudio.com/items?itemName=AdamWulkiewicz.UndoCloseDocument)
[![Rating](https://vsmarketplacebadges.dev/rating-short/AdamWulkiewicz.UndoCloseDocument.svg)](https://marketplace.visualstudio.com/items?itemName=AdamWulkiewicz.UndoCloseDocument)
![License](https://img.shields.io/github/license/awulkiew/undo-close-document.svg)
[![Donate](https://img.shields.io/badge/Donate-_-yellow.svg)](https://awulkiew.github.io/donate)

This extension allows to reopen recently closed document tab.

1. **Right-click** one of the opened tabs or go to menu **Window**
2. Choose **Undo Close Document**

![Exclude From Build](images/preview.png)

Keyboard shortcut can be set in **Tools -> Options -> Environment -> Keyboard** command **Window.UndoCloseDocument**.

Maximum numbers of documents remembered and shown on the list can be set in **Tools -> Options -> Undo Close Document -> General**.

This extension has to be autoloaded with Visual Studio to work properly. It supports Visual Studio as old as 2013 which means it's loaded synchroniously. In Visual Studio 2019 synchronious autoloading is disabled by default. If you want to use this extension with Visual Studio 2019 you have to **Allow synchronous autoload of extensions** in **Tools -> Options -> Environment -> Extensions**.

You can download this extension from [Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=AdamWulkiewicz.UndoCloseDocument) or [GitHub](https://github.com/awulkiew/undo-close-document/releases).
