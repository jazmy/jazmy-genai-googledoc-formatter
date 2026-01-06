# Google Docs Infographic Generator & Book Formatter

A Google Apps Script that combines AI-powered infographic generation with professional document formatting. 

**Read our full blog article for more details:** [GenAI Secret Sauce for Google Docs Professional Formatting and Infographics](https://genaisecretsauce.com/genai-secret-sauce-for-google-docs-professional-formatting-and-infographics)


## Features

- **Infographic Generation**: Uses Google Gemini's Nano Banana Pro model to analyze document content and generate relevant infographics
- **Document Formatting**: Professional styling for headings, body text, lists, tables, links, and images
- **Customizable**: All colors, fonts, sizes, and branding configurable at the top of the script

## Setup

1. Open your Google Doc
2. Go to **Extensions → Apps Script**
3. Paste the contents of `InfographicGenerator.gs`
4. Save and refresh your Google Doc
5. Two new menus will appear: **🎨 Infographic Generator** and **📄 Book Formatting**

### API Key Setup (for Infographics)

1. Get a free Gemini API key at https://aistudio.google.com/apikey
2. In Google Docs: **🎨 Infographic Generator → Set API Key**
3. Enter your API key when prompted

---

## Menu Reference

### 🎨 Infographic Generator

| Menu Item | Function | Description |
|-----------|----------|-------------|
| **Generate →** | | |
| Smart Analysis (Recommended) | `generateSmartInfographics` | Analyzes entire document holistically, determines optimal infographic placement (document-wide, multi-section, or single-section), then generates and embeds all |
| Per-Section Only | `generatePerSectionInfographics` | Analyzes each section independently, generates infographics only for individual sections that need visuals |
| Document Summary Only | `generateDocumentSummary` | Creates a single executive overview infographic at the beginning of the document |
| **Analyze Document (Preview)** | `analyzeDocumentPreview` | Shows what infographics would be generated without actually creating them - lets you review the plan first |
| **Set API Key** | `showApiKeyDialog` | Prompts you to enter your Gemini API key (stored in script properties) |
| **Help** | `showInfographicHelp` | Displays help information about the infographic generator |

### 📄 Book Formatting

| Menu Item | Function | Description |
|-----------|----------|-------------|
| **Format Everything (Fresh)** | `formatAllFresh` | Full document format - applies base font globally first, then formats all elements. Fastest for first-time use |
| **Format Everything (Resume)** | `formatAllResume` | Checks each element before formatting, skips already-formatted ones. Use if interrupted mid-process |
| **Individual Formatters →** | | |
| Format H1 Headings | `formatH1Only` | Adds red vertical bar prefix (┃), primary font 22pt, bold, dark text with red bar |
| Format H2 Headings | `formatH2Only` | Adds arrow prefix (→), primary font 20pt, bold, red text |
| Format H3 Headings | `formatH3Only` | Primary font 18pt, bold, dark gray text |
| Format H4 Headings | `formatH4Only` | Secondary font 16pt, normal weight (not bold), dark gray |
| Format H5 Headings | `formatH5Only` | Primary font 14pt, bold, dark gray |
| Format Body Text | `formatBodyOnly` | Primary font 11pt, charcoal color, 1.15 line spacing |
| Format Lists | `formatListsOnly` | Compact spacing (0 before/after), auto-bolds text before colons or dashes |
| Format Tables | `formatTablesOnly` | Dark header row, alternating row colors, proportional column widths, subtle borders |
| Format Links | `restoreLinkFormatting` | Restores blue color (#1155CC) and underline to all hyperlinks |
| **Utilities →** | | |
| Resize Images to Full Width | `resizeImagesOnly` | Scales all images to 1200pt width while maintaining aspect ratio |
| Bold List Terms | `boldListTermsOnly` | Bolds text before `:` or ` - ` in list items (for definition-style lists) |
| Promote Headings (H2→H1, etc.) | `promoteHeadings` | Moves all headings up one level (H2→H1, H3→H2, H4→H3, H5→H4) |
| Remove Horizontal Lines | `removeHorizontalLines` | Deletes all horizontal rules and dash/underscore lines |
| Remove HTML Comments | `cleanupComments` | Removes paragraphs that are `<!-- comment -->` format |
| Remove Bookmarks | `removeAllBookmarks` | Deletes all bookmarks from the document |
| **Preview (First 5) →** | | |
| Preview All Formatting | `formatFirst5All` | Formats only the first 5 of each element type - lets you check styling before full run |
| Preview Tables Only | `formatFirst5Tables` | Formats only the first 5 tables |
| **Help** | `showFormattingHelp` | Displays help information about book formatting options and current font settings |

---

## Configuration

All settings are at the top of the script for easy customization.

### CONFIG (API & Processing)

```javascript
const CONFIG = {
  GEMINI_API_KEY_PROPERTY: 'GEMINI_API_KEY',
  TEXT_MODEL: 'gemini-3-flash-preview',
  IMAGE_MODEL: 'gemini-3-pro-image-preview',
  API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/models',
  MIN_SECTION_LENGTH: 100,    // Minimum characters to consider for infographic
  MAX_SECTIONS_TO_PROCESS: 10 // Maximum infographics to generate
};
```

### BRANDING (Logo)

```javascript
const BRANDING = {
  INCLUDE_LOGO: true,  // Set to false to disable logo
  LOGO_URL: 'https://example.com/logo.png',
  LOGO_POSITION: 'bottom-right corner'
};
```

### COLORS (Infographic Palette)

```javascript
const COLORS = {
  // Backgrounds
  PRIMARY_BACKGROUND: '#FFFFFF',
  SECONDARY_BACKGROUND: '#F8FAFC',
  SUBTLE_BACKGROUND: '#E2E8F0',

  // Primary palette
  PRIMARY_ACCENT: '#B4007D',    // Deep magenta
  SECONDARY_ACCENT: '#7B2D8E', // Rich purple
  EMPHASIS: '#E50914',          // Red
  CONTRAST: '#1E3A8A',          // Deep blue
  SUCCESS: '#F59E0B',           // Golden yellow

  // Text
  TEXT_PRIMARY: '#1E293B',
  TEXT_SECONDARY: '#666666'
};
```

### FONTS (Primary Font Configuration)

```javascript
const FONTS = {
  PRIMARY: 'Google Sans Text',    // Main font for headings and body
  SECONDARY: 'Montserrat',        // Accent font (used for H4)
  CODE: 'Roboto Mono',            // Monospace font for code
};
```

### STYLES (Document Formatting)

```javascript
const STYLES = {
  fonts: {
    h1: FONTS.PRIMARY,
    h2: FONTS.PRIMARY,
    h3: FONTS.PRIMARY,
    h4: FONTS.SECONDARY,
    h5: FONTS.PRIMARY,
    heading: FONTS.PRIMARY,
    body: FONTS.PRIMARY,
    code: FONTS.CODE,
  },
  sizes: {
    h1: 22,
    h2: 20,
    h3: 18,
    h4: 16,
    h5: 14,
    body: 11,
    code: 10,
  },
  colors: {
    headingDark: '#444444',
    headingRed: '#980000',
    charcoal: '#221F1F',
    darkGray: '#333333',
    darkGray3: '#999999',
    white: '#FFFFFF',
    tableHeader: '#333333',
    tableHeaderText: '#FFFFFF',
    tableBorder: '#E8E8E8',
    tableAltRow: '#F9F9F9',
    linkBlue: '#1155CC',
  },
  spacing: {
    h1Before: 24,
    h1After: 12,
    h2Before: 18,
    h2After: 8,
    h3Before: 14,
    h3After: 6,
    paragraphAfter: 8,
    lineSpacing: 1.15,
  },
  imageWidth: 1200,  // Used by "Resize Images to Full Width" utility
  infographicWidth: 468,  // Full page width for generated infographics
};
```

---

## Typical Workflows

### First-Time Document Formatting

1. Run **Format Everything (Fresh)**
2. Review the document
3. Use Individual Formatters to adjust specific elements if needed

### Interrupted Formatting

1. Run **Format Everything (Resume)** - skips already-formatted elements

### Testing Before Full Format

1. Run **Preview → Preview All Formatting** to format first 5 of each element
2. Scroll through to check styling
3. If satisfied, run **Format Everything (Fresh)**

### Generating Infographics

1. Set your Gemini API key (one-time)
2. Run **Analyze Document (Preview)** to see recommendations
3. Run **Generate → Smart Analysis** to create and embed infographics

---

## Infographic Types

The AI analyzes content and generates appropriate infographic types:

- **Flowcharts** - Process flows and workflows
- **Statistics** - Data and numbers visualization
- **Comparisons** - Side-by-side comparisons
- **Timelines** - Sequential information
- **Hierarchies** - Organizational structures
- **Concept Maps** - Related ideas and concepts
- **Overview** - Document-wide summaries

---

## Infographic Scopes

| Scope | Description | Placement |
|-------|-------------|-----------|
| **Document-wide** | Executive summary for the entire document | Beginning of document |
| **Multi-section** | Spans 2+ related sections | Above first relevant section |
| **Single-section** | Standalone content visualization | Above that section |

---

## Notes

- **Custom fonts**: If a specified font is not available, Google Docs will use a fallback font. By default, the primary font is Google Sans Text and secondary font is Montserrat.
- **API Limits**: The script includes delays between API calls to avoid rate limiting
- **Large Documents**: Use Resume mode if formatting times out on very large documents
- **Logo in Infographics**: The logo is sent as a reference image to Gemini; reproduction accuracy may vary
- **Infographic Sizing**: Generated infographics are automatically sized to full page width (468pt = 6.5 inches, standard Google Docs content width) with 16:9 aspect ratio
