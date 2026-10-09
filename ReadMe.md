---
title: "2htm ReadMe"
author: "Jamal Mazrui"
---

# 2htm ReadMe

2htm converts Microsoft Word, Excel, PowerPoint, PDF, Markdown and other files to accessible HTML, or to plain text. Headings, lists, tables and image descriptions come through in a structure a screen reader can move around in.

This is the quick start. The full guide is `help\2htm.htm`.

## Install

1. Download `2htm_setup.exe` from the [2htm releases page](https://github.com/jamalmazrui/2htm/releases).
2. Run it. It asks for administrator rights, because it installs for everyone on the computer.
3. On the last page, leave **Launch 2htm now** checked and press Finish. Read the results box, then close it; 2htm opens.

You need Windows 10 or 11, 64-bit, and the Office program for each kind of Office file you convert: Word for `.docx`, Excel for `.xlsx`, PowerPoint for `.pptx`. Markdown needs nothing more.

## Convert a file

1. Press **Alt+Control+2** from anywhere in Windows (except inside Word, which keeps that key for Heading 2). The 2htm dialog opens with focus in **Source files**.
2. Type a file name, a folder, or a pattern such as `*.docx`.
3. Press Enter.

A results box says what was converted. Each result is named after its file, such as `report.htm`.

Or, in File Explorer, press **Shift+F10** on a file and then **2**, for Convert via 2htm. The result is written next to the file.

## From the command line

In a Command Prompt in the program folder:

```cmd
2htm report.docx
2htm *.md -o html
2htm -h
```

## Keys in the dialog

- **Alt** with an underlined letter moves to that control.
- **Enter** starts; **Escape** closes the dialog.
- **F1** shows Help.
- **F11** checks the web for a newer version of 2htm and offers to install it.

All the keys are listed in `help\Hotkeys.htm`.

## Learn by listening

Ten short spoken walks teach 2htm, each a few minutes long, in two voices: a host, and a screen reader saying what you would hear. They are installed in the `help\tutorials` folder as mp3 files, with a playlist, and as text in `Tutorials.htm`. Start with walk 0, the overview.

## When something goes wrong

Every run keeps a log in `%LOCALAPPDATA%\2htm\logs`, one file per run. Zip that folder and send it with a description of what happened.

## License

2htm is free and open source under the MIT License. See `License.htm`.
