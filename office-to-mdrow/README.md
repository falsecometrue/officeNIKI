# office to mdraw

Convert Microsoft Office files from the VS Code Explorer context menu.

The extension currently converts Excel and Word files into editable text-based formats. Conversion runs locally and writes the generated file next to the selected Office file.

## Supported Conversions

| Input | Output | Status |
|---|---|---|
| `.xlsx` | `.drawio` | Supported |
| `.xlsx` | `.md` and `resources/` | Supported |
| `.docx` | `.md` and `resources/` | Supported |
| `.pptx` | Marp Markdown and `.drawio.svg` slides | Supported |

## How to Use

1. Open a folder in VS Code.
2. Right-click an Office file in the Explorer.
3. Select one of the conversion commands:

   - `Convert Excel to Draw.io`

     ![Convert Excel to Draw.io context menu](https://raw.githubusercontent.com/comecomenakau/officeNIKI/30.developAndTest/office-to-mdrow/image/README/1778455217913.png)

   - `Convert Excel to Markdown`

   - `Convert Word to Markdown`

     ![Convert Word to Markdown context menu](https://raw.githubusercontent.com/comecomenakau/officeNIKI/30.developAndTest/office-to-mdrow/image/README/1778455291868.png)

   - `Convert PowerPoint to Marp`

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

PowerPoint `.pptx` to Marp Markdown conversion treats Marp as the display format and Draw.io SVG as the editable slide source.

Output:

```text
pptx-file-name/
  pptx-file-name.md
  slide/
    slide001.drawio.svg
    slide002.drawio.svg
```

Marp structure:

```md
---
marp: true
---

![bg contain](slide/slide001.drawio.svg)

---

![bg contain](slide/slide002.drawio.svg)
```

- Converts each PowerPoint slide into one `.drawio.svg` file.
- Embeds slide images, backgrounds, shapes, and tables into that slide-level SVG.
- Embeds Draw.io editing data in the SVG so the SVG itself is both the display asset and the editable source.
- Does not generate separate `.drawio` files for PowerPoint slides.
- Does not output separate slide image/background files for PowerPoint Marp.
- Uses Marp only as a presentation/display wrapper that references the `.drawio.svg` files.
- Avoids duplicated editable/display artifacts, so there is no risk of `.drawio` and exported `.svg` drifting out of sync.

## Requirements

No Python installation is required. The conversion logic is implemented in TypeScript and runs inside the VS Code extension environment.

The extension is disabled in untrusted workspaces because conversion reads Office package contents and writes generated files next to the selected source file.

Very large or suspicious Office packages are rejected before conversion, including packages with too many ZIP entries, oversized XML entries, oversized images, or excessive sheet/slide counts.

## Data Usage

- All conversion runs locally on your machine.
- Selected Office file contents are used only to generate the converted output.
- This extension does not upload Office files or conversion results to any external server.
- This extension does not collect telemetry.

## Output Policy

- If a conversion target already exists, the extension asks before overwriting it.
- Intermediate JSON is not kept as a final output.
- Word images are written to `resources/`.
- Excel Draw.io images are embedded in the Draw.io XML.
- Excel Markdown images are written to `resources/`.
- PowerPoint Marp output references editable `.drawio.svg` slide files only, without separate `.drawio` files.
- Embedded images are limited to common raster web formats: PNG, JPEG, GIF, and WebP.

## Maintenance

Dependency updates are checked weekly by Dependabot for npm packages under `office-to-mdrow/` and GitHub Actions workflows. Pull requests and pushes that touch this extension run CI with install, compile, tests, dependency audit, VSIX packaging, and VSIX content listing.

## Release Checklist

Run the same checks locally before publishing a VSIX:

```sh
cd office-to-mdrow
npm run release:check
```

The release check performs a clean install, compiles TypeScript, runs unit and security tests, runs runtime and full dependency audits, builds the VSIX, and lists the files included in the package.
