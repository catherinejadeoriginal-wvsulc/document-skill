# Skill: DOCX Document Generation

Generate professional .docx files using the `docx` npm package.

## Prerequisites

```bash
npm install docx
```

## Core Pattern

```javascript
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, Row, Cell } = require('docx');

const doc = new Document();
// ... build document ...

Packer.toBuffer(doc).then((buffer) => {
  // Write buffer to file or stream
});
```

## Common Operations

### Text & Paragraphs
```javascript
const p = doc.addParagraph("Hello", {
  heading: HeadingLevel.HEADING_1
});
p.addRun("bold text").bold = true;
p.addRun(" normal text").fontSize = 12;
p.addRun(" colored text").color = "000000";
p.alignment = "center";
```

### Tables
```javascript
const table = doc.addTable(3, 2);
table.setWidth("100%");
table.alignment = "center";
table.getCell(0, 0).addParagraph("Header");
// merge cells
table.getCell(0, 0).merge(table.getCell(0, 1));
```

### Images
```javascript
// Note: Image support requires additional processing
// For base64 images or file paths, use a different approach
```

### Page Setup
```javascript
const section = doc.sections[0];
section.setPageSize({
  width: 11905.6, // 8.5 inches in twips
  height: 16838.4, // 11 inches in twips
});
section.orientation = "portrait";
section.setMargins({
  top: 1440, // 1 inch in twips
  bottom: 1440,
  left: 1440,
  right: 1440,
});
```

### Headers & Footers
```javascript
const header = doc.createHeader();
header.addParagraph("Header Text", {
  alignment: "center"
});

const footer = doc.createFooter();
const footerPara = footer.addParagraph("Page ");
footerPara.addRun("1").pageNumber();
```

### Styles
```javascript
// Styles are applied directly to elements
// For custom styles, use the styles API
```

### Page Breaks
```javascript
doc.addSectionBreak();
```

### Lists
```javascript
doc.addParagraph("Item 1", {
  bullet: {
    level: 0
  }
});

doc.addParagraph("Item 2", {
  numbering: {
    reference: "1",
    level: 0
  }
});
```

## Workflow

1. Load skill: `/skill docx` to activate this skill
2. Define document structure (sections, styles, content)
3. Generate the .docx using the docx npm package
4. Always run the script and verify the output
