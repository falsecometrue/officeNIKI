# office to mdraw

Convert Microsoft Office files from the VS Code Explorer context menu.

The extension currently converts Excel and Word files into editable text-based formats. Conversion runs locally and writes the generated file next to the selected Office file.

## Supported Conversions

| Input | Output | Status |
|---|---|---|
| `.xlsx` | `.drawio` | Supported |
| `.xlsx` | `.md` and `resources/` | Supported |
| `.docx` | `.md` and `resources/` | Supported |
| `.pptx` | Marp Markdown and `.drawio.svg` resources | Not implemented yet. Planned for a future release. |

## How to Use

1. Open a folder in VS Code.
2. Right-click an Office file in the Explorer.
3. Select one of the conversion commands:

   - `Convert Excel to Draw.io`

     ![Convert Excel to Draw.io context menu](https://raw.githubusercontent.com/comecomenakau/officeNIKI/30.developAndTest/office-to-mdrow/image/README/1778455217913.png)

   - `Convert Excel to Markdown`

   - `Convert Word to Markdown`

     ![Convert Word to Markdown context menu](https://raw.githubusercontent.com/comecomenakau/officeNIKI/30.developAndTest/office-to-mdrow/image/README/1778455291868.png)

4. The converted file is created in the same folder as the source file.

   Excel to Draw.io result:

   ![Excel to Draw.io conversion result](https://raw.githubusercontent.com/comecomenakau/officeNIKI/30.developAndTest/office-to-mdrow/image/README/1778455462127.png)

   Word to Markdown result:

   ![Word to Markdown conversion result](https://raw.githubusercontent.com/comecomenakau/officeNIKI/30.developAndTest/office-to-mdrow/image/README/1778455542849.png)

## Conversion Details

### Excel to Draw.io

- Reads `.xlsx` files as Office Open XML.
- Converts sheets into Draw.io pages.
- Converts cell text, basic tables, shapes, connectors, and embedded images.
- Embeds Excel images as base64 data URIs inside the Draw.io XML.
- Outputs a `.drawio` file next to the source workbook.

### Excel to Markdown

- Reads `.xlsx` files as Office Open XML.
- Outputs a folder named after the workbook.
- Writes one Markdown file named after the workbook.
- Groups each sheet under a `# シート名` heading.
- Exports sheet images into `resources/` as `シート名-1.png`, `シート名-2.png`, and so on.
- References images from Markdown using relative paths.

### Word to Markdown

- Reads `.docx` files as Office Open XML.
- Converts headings, paragraphs, tables, images, and simple shape text.
- Exports Word images into a `resources/` folder.
- References images from Markdown using relative paths.
- Outputs a `.md` file next to the source document.

### PowerPoint to Marp

PowerPoint `.pptx` to Marp Markdown conversion is not implemented yet. The planned design treats Marp as the display format and Draw.io SVG as the editable slide source.

Planned output:

```text
pptx-file-name/
  pptx-file-name.md
  slide/
    slide001.drawio.svg
    slide002.drawio.svg
```

Planned Marp structure:

```md
---
marp: true
---

![bg contain](slide/slide001.drawio.svg)

---

![bg contain](slide/slide002.drawio.svg)
```

Design policy:

- Converts each PowerPoint slide into one `.drawio.svg` file.
- Embeds slide images, backgrounds, shapes, and tables into that slide-level SVG.
- Embeds Draw.io editing data in the SVG so the SVG itself is both the display asset and the editable source.
- Does not generate separate `.drawio` files for PowerPoint slides.
- Does not output separate slide image/background files for PowerPoint Marp.
- Uses Marp only as a presentation/display wrapper that references the `.drawio.svg` files.
- Avoids duplicated editable/display artifacts, so there is no risk of `.drawio` and exported `.svg` drifting out of sync.

## Requirements

No Python installation is required. The conversion logic is implemented in TypeScript and runs inside the VS Code extension environment.

## Data Usage

- All conversion runs locally on your machine.
- Selected Office file contents are used only to generate the converted output.
- This extension does not upload Office files or conversion results to any external server.

## Output Policy

- Intermediate JSON is not kept as a final output.
- Word images are written to `resources/`.
- Excel Draw.io images are embedded in the Draw.io XML.
- Excel Markdown images are written to `resources/`.
- PowerPoint Marp output will reference editable `.drawio.svg` slide resources only, without separate `.drawio` files.
