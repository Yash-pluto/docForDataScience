# 📄 Docx Exam Strategy Generator

A minimal TypeScript project that programmatically generates a beautifully formatted Microsoft Word (.docx) study guide using the docx library.

> This specific script generates a Data Science Exam Strategy cheat sheet with custom typography, shaded tables, and clean bullet points.

## ✨ Features

- TypeScript Ready: Fully typed setup for easy editing and autocompletion.

- Custom Styling: Programmatically sets fonts (Calibri), borders, table shading, and heading alignments.

- Zero-to-Docx: Bypasses manual Word formatting to generate a ready-to-print study guide instantly.

## 🚀 Quick Start

1. Prerequisites
   Make sure you have Node.js installed on your machine.

2. Install Dependencies
   Initialize your project and install the required packages (if you haven't already):

```bash
npm init -y
npm install docx
npm install -D @types/node
```

### Tip: To run the file directly without compiling to JS first, install tsx globally:

```bash
npm install -g tsx
```

3. Generate the Document
   Run the TypeScript file directly:

```bash
tsx generate_doc.ts
```

(Note: Replace generate_doc.ts with your actual file name).

You will see a confirmation message in the console, and a new DS_Exam_Strategy.docx file will appear in your project folder. Open it up and enjoy!

### 🛠️ Customization

- Want to make a study guide for a different subject? Open the .ts file and tweak the text inside the heading1, heading2, para, and makeTable function calls.

- You can also change the MAIN_FONT variable to switch up the typography, or adjust the color hex codes for a different aesthetic.
