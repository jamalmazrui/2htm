# 2htm — Convert Office, PDF, and Markdown Files to Accessible HTML

I am pleased to share a new free, open-source tool I have just released on GitHub. `2htm` is a single, independent executable that runs on any modern Windows computer with Microsoft Office installed — no installer, no runtime to download, no sidecar DLLs.

## 2htm

<https://github.com/jamalmazrui/2htm>

`2htm` converts Microsoft Word, Excel, PowerPoint, PDF, HTML, Markdown, JSON, CSV, and plain text files into a single-file HTML equivalent that opens in any modern browser — Chrome, Edge, Firefox, Safari — with no companion folders or special viewer required. It can also produce plain-text output instead.

The same executable runs in two modes with the same set of options. Command-line mode handles scripted and batch use. GUI mode opens a small, keyboard-accessible dialog with every option exposed as a field or checkbox, for users who prefer an interactive workflow. A single flag (`-g`) toggles between them.

The conversion aims for WCAG 2.2 AAA conformance to the extent the source document's structure and content permit. Landmarks, headings, table markup, alt-text propagation, color contrast, and language declaration are preserved or inferred where possible. The output is designed to work cleanly with screen readers and to reflow on small screens.

`2htm` fits naturally into pipelines that produce "alternate formats" — accessible versions of public documents for users with disabilities. Call it synchronously from a batch file, a scheduled task, or a CI job; inspect the exit code; serve the resulting `.htm` files directly from any web host. Because the output is self-contained, downstream steps never have to track sidecar resources.

The tool is released under the permissive MIT license, so programmers can freely modify or extend the code. The entire source is one C# file (`2htm.cs`); C# is Microsoft's flagship application language, well-documented, with a mature free toolchain from Microsoft (Visual Studio Community or Visual Studio Build Tools, both free downloads) for developers who want to build from source.

This project was developed in collaboration with Anthropic's Claude AI assistant. Feedback and contributions are welcome.

\#accessibility \#wcag \#a11y \#assistivetechnology \#commandlinetools \#opensource
