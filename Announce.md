# 2htm — Produce accessible HTML versions of popular file formats from a single Windows executable

I am pleased to share a new free, open-source tool I have just released on GitHub. `2htm` is a single, independent executable that runs on any modern Windows computer with Microsoft Office installed (no installer to run, no runtime to download, no additional DLLs to track).

## 2htm

<https://github.com/jamalmazrui/2htm>

`2htm` converts Microsoft Word, Excel, PowerPoint, PDF, and Markdown files into a single-file, HTML equivalent that opens in any modern browser — Chrome, Edge, Firefox, Safari — with no companion folders or special viewer required. It can also produce plain-text output instead.

The single executable can run in two modes with the same set of options. Command-line mode handles scripted and batch use. GUI mode opens a small, keyboard-accessible dialog with every option exposed as a field or checkbox, for users who prefer an interactive workflow. 

The conversion aims for WCAG 2.2 AAA conformance to the extent the source document's structure and content permit. Landmarks, headings, table markup, alt-text propagation, color contrast, and language declaration are preserved or inferred where possible. The output is designed to work cleanly with screen readers and to reflow on small screens.

`2htm` fits naturally into pipelines that produce "alternate formats" — accessible versions of public documents for users with disabilities. Call it synchronously from a batch file or scheduled task -- even serve the resulting `.htm` files directly from any web host. 

The tool is released under the permissive MIT license, so programmers can freely modify or extend the code. It is written in C# -- Microsoft's flagship application language, well-documented, with a mature free toolchain.

This project was developed in collaboration with Anthropic's Claude AI.

You can download the whole project in a single zip archive using the following link:

<http://GitHub.com/JamalMazrui/2htm/archive/main.zip>
