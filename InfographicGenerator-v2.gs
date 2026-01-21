/**
 * Infographic Generator for Google Docs V2
 *
 * Analyzes document content and generates infographics using
 * Google Gemini's Nano Banana Pro (gemini-3-pro-image-preview) model.
 * Update in V2 - for longer docs allows you to generate 10 infographics at a time and resume so the script does not time out.
 * 
 * Setup:
 * 1. Get a Gemini API key from https://aistudio.google.com/apikey
 * 2. Set your API key in Script Properties: GEMINI_API_KEY
 */

// ============================================================
// CONFIGURATION - Adjust these settings to customize behavior
// ============================================================

const CONFIG = {
  // API Settings
  GEMINI_API_KEY_PROPERTY: 'GEMINI_API_KEY',
  TEXT_MODEL: 'gemini-3-flash-preview',
  IMAGE_MODEL: 'gemini-3-pro-image-preview',
  API_BASE_URL: 'https://generativelanguage.googleapis.com/v1beta/models',

  // Processing Limits
  MIN_SECTION_LENGTH: 100,    // Minimum characters to consider for infographic
  BATCH_SIZE: 10,             // Sections to process per run (to avoid timeout)
  MAX_TOTAL_INFOGRAPHICS: 50  // Absolute max infographics for any document
};

// ============================================================
// BRANDING - Logo settings
// ============================================================

const BRANDING = {
  INCLUDE_LOGO: true,  // Set to false to disable logo
  LOGO_URL: 'https://genaisecretsauce.com/content/images/2024/11/secret_sauce_logo_800px-3.png',
  LOGO_POSITION: 'bottom-right corner'
};

// ============================================================
// COLOR PALETTE - Customize infographic colors
// ============================================================

const COLORS = {
  // Backgrounds
  PRIMARY_BACKGROUND: '#FFFFFF',    // Crisp white
  SECONDARY_BACKGROUND: '#F8FAFC',  // Soft off-white
  SUBTLE_BACKGROUND: '#E2E8F0',     // Soft grey

  // Primary palette
  PRIMARY_ACCENT: '#B4007D',        // Deep magenta - primary accents and connectors
  SECONDARY_ACCENT: '#7B2D8E',      // Rich purple - secondary elements
  EMPHASIS: '#E50914',              // Red - emphasis and key highlights
  CONTRAST: '#1E3A8A',              // Deep blue - contrast elements
  SUCCESS: '#F59E0B',               // Golden yellow - positive/success indicators

  // Text
  TEXT_PRIMARY: '#1E293B',          // Dark slate - all text
  TEXT_SECONDARY: '#666666'         // Grey - captions and secondary text
};

// ============================================================
// DOCUMENT FORMATTING STYLES
// ============================================================

// Primary fonts - change these to customize document fonts
const FONTS = {
  PRIMARY: 'Google Sans Text',    // Main font for headings and body
  SECONDARY: 'Montserrat',        // Accent font (used for H4)
  CODE: 'Roboto Mono',            // Monospace font for code
};

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
  imageWidth: 1200,
};

// ============================================================
// INFOGRAPHIC STYLE TEMPLATE - Uses variables from above
// ============================================================

const INFOGRAPHIC_STYLE = `
Create a professional tech infographic

STYLE: Clean modern tech presentation on white background.
Crisp white (${COLORS.PRIMARY_BACKGROUND}) or soft off-white (${COLORS.SECONDARY_BACKGROUND}) background
with subtle light grey geometric grid pattern. High-fidelity
3D rendered objects with soft drop shadows and subtle reflections.
Premium corporate technology aesthetic, airy and sophisticated.

COLORS: Use the following colors on a white background:
- Primary accent (${COLORS.PRIMARY_ACCENT}) for primary accents and connectors
- Secondary accent (${COLORS.SECONDARY_ACCENT}) for secondary elements
- Emphasis color (${COLORS.EMPHASIS}) for emphasis and key highlights
- Contrast color (${COLORS.CONTRAST}) for contrast elements and alternative accents
- Success color (${COLORS.SUCCESS}) for positive/success indicators
- Text color (${COLORS.TEXT_PRIMARY}) for all text
- Subtle background (${COLORS.SUBTLE_BACKGROUND}) for subtle backgrounds
Mix these colors harmoniously throughout the infographic.

TYPOGRAPHY: Title in a rounded rectangular box with slightly
translucent light grey/white background, floating at the top
center of the infographic. Bold ${COLORS.TEXT_PRIMARY} sans-serif text.
Clean, modern tech font (Inter/SF Pro style). All text highly
legible with strong contrast against light backgrounds.
Section headers in ${COLORS.PRIMARY_ACCENT} or ${COLORS.SECONDARY_ACCENT}.

VISUALS: Key concepts illustrated as polished 3D rendered objects
with soft diffused shadows (no harsh edges). Subtle ambient occlusion.
Clean vector connector lines connecting related elements.
Floating information boxes with soft shadows and colored left borders.
NO neon glow effects — use refined shadows and subtle gradients instead.

LAYOUT: Title banner spanning full width at top with optional
light grey accent bar. Clear visual zones with generous whitespace.
Logical left-to-right or top-to-bottom flow. Clean thin-line
connectors and arrows showing relationships.

BRANDING: ${BRANDING.INCLUDE_LOGO ? `A reference image of the logo is attached.
Include this logo small in the ${BRANDING.LOGO_POSITION} of the infographic.
The logo must match the image
Keep it minimal and tasteful - do not overuse the branding.` : 'No branding required.'}

FORMAT: 16:9 aspect ratio, presentation slide compatible.
`;

/**
 * Adds custom menus to Google Docs
 */
function onOpen() {
  const ui = DocumentApp.getUi();

  // Infographic Generator Menu
  ui.createMenu('🎨 Infographic Generator')
    .addItem('▶️ Generate Infographics', 'generatePerSectionInfographics')
    .addItem('📊 Analyze Document (Preview)', 'analyzeDocumentPreview')
    .addSeparator()
    .addSubMenu(ui.createMenu('Other Options')
      .addItem('Smart Analysis (Experimental)', 'generateSmartInfographics')
      .addItem('Document Summary Only', 'generateDocumentSummary'))
    .addSeparator()
    .addItem('🔑 Set API Key', 'showApiKeyDialog')
    .addItem('❓ Help', 'showInfographicHelp')
    .addToUi();

  // Document Formatting Menu
  ui.createMenu('📄 Book Formatting')
    .addItem('Format Everything (Fresh)', 'formatAllFresh')
    .addItem('Format Everything (Resume)', 'formatAllResume')
    .addSeparator()
    .addSubMenu(ui.createMenu('Individual Formatters')
      .addItem('Format H1 Headings', 'formatH1Only')
      .addItem('Format H2 Headings', 'formatH2Only')
      .addItem('Format H3 Headings', 'formatH3Only')
      .addItem('Format H4 Headings', 'formatH4Only')
      .addItem('Format H5 Headings', 'formatH5Only')
      .addItem('Format Body Text', 'formatBodyOnly')
      .addItem('Format Lists', 'formatListsOnly')
      .addItem('Format Tables', 'formatTablesOnly')
      .addItem('Format Links', 'restoreLinkFormatting'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Utilities')
      .addItem('Resize Images to Full Width', 'resizeImagesOnly')
      .addItem('Bold List Terms', 'boldListTermsOnly')
      .addItem('Promote Headings (H2→H1, etc.)', 'promoteHeadings')
      .addItem('Remove Horizontal Lines', 'removeHorizontalLines')
      .addItem('Remove HTML Comments', 'cleanupComments')
      .addItem('Remove Bookmarks', 'removeAllBookmarks'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Preview (First 5)')
      .addItem('Preview All Formatting', 'formatFirst5All')
      .addItem('Preview Tables Only', 'formatFirst5Tables'))
    .addSeparator()
    .addItem('Help', 'showFormattingHelp')
    .addToUi();
}

/**
 * Smart analysis - determines optimal infographic placement (single, multi-section, or document-wide)
 * Now with incremental processing and resume capability
 */
function generateSmartInfographics() {
  const ui = DocumentApp.getUi();
  const docId = DocumentApp.getActiveDocument().getId();

  const apiKey = getApiKey();
  if (!apiKey) {
    ui.alert('API Key Required',
      'Please set your Gemini API key first.\n\nGo to: Infographic Generator → Set API Key',
      ui.ButtonSet.OK);
    return;
  }

  // Check for existing progress
  const progress = getSmartProgress(docId);
  let resumeMode = false;
  let savedPlan = null;
  let startIndex = 0;
  
  if (progress && progress.plan && progress.plan.infographics) {
    const response = ui.alert('Resume Previous Run?',
      `Found previous progress: ${progress.success} completed out of ${progress.plan.infographics.length} planned.\n\n` +
      'YES = Resume from where you left off\n' +
      'NO = Start fresh with new analysis',
      ui.ButtonSet.YES_NO);
    
    if (response === ui.Button.YES) {
      resumeMode = true;
      savedPlan = progress.plan;
      startIndex = progress.lastProcessedIndex + 1;
      log(`Resuming smart analysis from infographic ${startIndex + 1}`);
    } else {
      clearSmartProgress(docId);
      log('Starting fresh smart analysis');
    }
  }

  ui.alert('Processing',
    resumeMode 
      ? `Resuming from infographic ${startIndex + 1}...\n\nImages will be generated incrementally.`
      : 'Performing holistic document analysis...\n\nThis identifies single-section, multi-section, and document-wide infographic opportunities.\n\nImages will be generated incrementally as each is analyzed.',
    ui.ButtonSet.OK);

  try {
    let doc = DocumentApp.openById(docId);
    const sections = extractSections(doc);
    const fullText = getFullDocumentText(doc);

    if (sections.length === 0 && fullText.length < CONFIG.MIN_SECTION_LENGTH) {
      ui.alert('No Content', 'Document has insufficient content for infographic generation.', ui.ButtonSet.OK);
      return;
    }

    // Get or create the infographic plan
    let infographicPlan;
    if (resumeMode && savedPlan) {
      infographicPlan = savedPlan;
      log(`Using saved plan with ${infographicPlan.infographics.length} infographics`);
    } else {
      // Perform holistic analysis
      log('Performing holistic document analysis...');
      infographicPlan = analyzeDocumentHolistically(sections, fullText, apiKey);
      log(`Infographic plan: ${JSON.stringify(infographicPlan)}`);
      
      // FALLBACK: If holistic analysis returns nothing, fall back to per-section
      if (!infographicPlan.infographics || infographicPlan.infographics.length === 0) {
        log('Holistic analysis returned no results. Falling back to per-section analysis...');
        
        const response = ui.alert('Smart Analysis Found Nothing',
          'The holistic analysis did not identify infographic opportunities.\n\n' +
          'Would you like to fall back to per-section analysis instead?',
          ui.ButtonSet.YES_NO);
        
        if (response === ui.Button.YES) {
          generatePerSectionInfographics();
          return;
        } else {
          ui.alert('Analysis Complete', 'No infographic opportunities identified.', ui.ButtonSet.OK);
          return;
        }
      }
    }

    // Process infographics incrementally
    const results = processInfographicPlanIncremental(docId, sections, infographicPlan, apiKey, startIndex);

    // Clear progress on successful completion
    if (results.completed) {
      clearSmartProgress(docId);
    }

    ui.alert(results.completed ? 'Complete' : 'Paused',
      `${results.completed ? 'Generated' : 'Progress so far:'} ${results.success} infographic(s).\n` +
      `• Document-wide: ${results.documentWide}\n` +
      `• Multi-section: ${results.multiSection}\n` +
      `• Single-section: ${results.singleSection}\n` +
      (results.failed > 0 ? `Failed: ${results.failed}\n` : '') +
      (!results.completed ? '\nRun again to continue.' : ''),
      ui.ButtonSet.OK);

  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    ui.alert('Error', 
      `An error occurred: ${error.message}\n\n` +
      `Progress has been saved. Run again to resume.`, 
      ui.ButtonSet.OK);
  }
}

/**
 * Generate infographics for individual sections
 * Processes incrementally with resume capability
 */
function generatePerSectionInfographics() {
  const ui = DocumentApp.getUi();
  const docId = DocumentApp.getActiveDocument().getId();

  const apiKey = getApiKey();
  if (!apiKey) {
    ui.alert('API Key Required',
      'Please set your Gemini API key first.\n\nGo to: 🎨 Infographic Generator → 🔑 Set API Key',
      ui.ButtonSet.OK);
    return;
  }

  // Get document sections first to show in dialog
  let doc = DocumentApp.openById(docId);
  const sections = extractSections(doc);
  const totalSections = sections.length;
  
  if (totalSections === 0) {
    ui.alert('No Content', 'No suitable sections found in the document.', ui.ButtonSet.OK);
    return;
  }

  // Check for existing progress
  const progress = getInfographicProgress(docId);
  let startIndex = 0;
  
  if (progress && progress.lastProcessedIndex >= 0) {
    // Show unified resume dialog
    startIndex = showResumeDialog(ui, progress, sections, docId);
    if (startIndex === -1) {
      return; // User cancelled
    }
  } else {
    // No progress - offer to start fresh or from specific section
    startIndex = showStartDialog(ui, sections, docId);
    if (startIndex === -1) {
      return; // User cancelled
    }
  }

  // Show brief processing message
  const startSection = sections[startIndex];
  const startLabel = startSection ? (startSection.heading || `Section ${startIndex + 1}`) : `Section ${startIndex + 1}`;
  
  ui.alert('Starting',
    `Beginning from: "${startLabel}"\n\n` +
    `Processing ${Math.min(CONFIG.BATCH_SIZE, totalSections - startIndex)} sections this run.\n` +
    `Images appear as each section completes.`,
    ui.ButtonSet.OK);

  try {
    let doc = DocumentApp.openById(docId);
    const sections = extractSections(doc);
    log(`Found ${sections.length} sections total`);

    if (sections.length === 0) {
      ui.alert('No Content', 'No suitable sections found for infographic generation.', ui.ButtonSet.OK);
      return;
    }

    if (startIndex >= sections.length) {
      ui.alert('Complete', 'All sections have already been processed. Clear progress to run again.', ui.ButtonSet.OK);
      clearInfographicProgress(docId);
      return;
    }

    // Process in batches to avoid timeout, but track overall progress
    const batchEndIndex = Math.min(startIndex + CONFIG.BATCH_SIZE, sections.length);
    const totalSections = sections.length;
    
    let success = progress ? progress.success : 0;
    let failed = progress ? progress.failed : 0;
    let skipped = progress ? progress.skipped : 0;
    let totalInfographicsGenerated = progress ? progress.totalGenerated : 0;

    log(`\nProcessing batch: sections ${startIndex + 1} to ${batchEndIndex} of ${totalSections}`);

    // Process sections incrementally - analyze and generate one at a time
    for (let i = startIndex; i < batchEndIndex; i++) {
      // Check if we've hit the absolute max
      if (totalInfographicsGenerated >= CONFIG.MAX_TOTAL_INFOGRAPHICS) {
        log(`Reached maximum infographic limit (${CONFIG.MAX_TOTAL_INFOGRAPHICS})`);
        break;
      }
      
      const section = sections[i];
      const sectionLabel = section.heading || `Section ${i + 1}`;
      
      log(`\n[${i + 1}/${totalSections}] Processing: ${sectionLabel}`);
      
      try {
        // Step 1: Analyze this section
        log(`  Analyzing section...`);
        const analysis = analyzeSection(section, apiKey);
        
        if (!analysis.needsInfographic) {
          log(`  Skipped: ${analysis.reason}`);
          skipped++;
          saveInfographicProgress(docId, i, success, failed, skipped, totalInfographicsGenerated, totalSections);
          continue;
        }
        
        log(`  Needs infographic: ${analysis.infographicType} - ${analysis.reason}`);
        
        // Step 2: Generate image immediately
        log(`  Generating infographic...`);
        const enrichedSection = {
          ...section,
          reason: analysis.reason,
          infographicType: analysis.infographicType,
          visualPrompt: analysis.visualPrompt
        };
        
        const imageBlob = generateInfographic(enrichedSection, apiKey);
        
        if (imageBlob) {
          // Step 3: Insert image immediately
          log(`  Inserting image into document...`);
          
          // Re-open document to get fresh state
          doc = DocumentApp.openById(docId);
          const body = doc.getBody();
          
          // Recalculate section positions (they may have shifted from previous inserts)
          const currentSections = extractSections(doc);
          if (i < currentSections.length) {
            insertImageAboveSection(body, currentSections[i], imageBlob);
          }
          
          doc.saveAndClose();
          
          success++;
          totalInfographicsGenerated++;
          log(`  ✓ Success! (${success} generated, ${totalInfographicsGenerated} total)`);
          
          // Delay to avoid rate limiting
          Utilities.sleep(2000);
        } else {
          failed++;
          log(`  ✗ Failed to generate image`);
        }
        
        // Save progress after each section
        saveInfographicProgress(docId, i, success, failed, skipped, totalInfographicsGenerated, totalSections);
        
      } catch (error) {
        log(`  ✗ Error: ${error.message}`);
        failed++;
        saveInfographicProgress(docId, i, success, failed, skipped, totalInfographicsGenerated, totalSections);
        
        // Continue to next section rather than aborting
        Utilities.sleep(1000);
      }
    }

    // Check if there are more sections to process
    const lastProcessedIndex = batchEndIndex - 1;
    const remainingSections = totalSections - batchEndIndex;
    const isComplete = remainingSections <= 0 || totalInfographicsGenerated >= CONFIG.MAX_TOTAL_INFOGRAPHICS;

    if (isComplete) {
      // All done! Clear progress
      clearInfographicProgress(docId);
      
      ui.alert('Complete',
        `✅ All sections processed!\n\n` +
        `✓ Generated: ${success} infographics\n` +
        `⊘ Skipped (no visual needed): ${skipped}\n` +
        (failed > 0 ? `✗ Failed: ${failed}\n` : '') +
        `\nTotal sections: ${totalSections}`,
        ui.ButtonSet.OK);
    } else {
      // More sections remain - keep progress saved
      ui.alert('Batch Complete - More Sections Remain',
        `📦 Batch complete! Processed sections ${startIndex + 1} to ${batchEndIndex}.\n\n` +
        `✓ Generated this batch: ${success - (progress ? progress.success : 0)}\n` +
        `✓ Total generated so far: ${success}\n` +
        `⊘ Skipped: ${skipped}\n` +
        (failed > 0 ? `✗ Failed: ${failed}\n` : '') +
        `\n📋 Remaining: ${remainingSections} sections\n\n` +
        `▶️ Run "Per-Section Only" again to continue.`,
        ui.ButtonSet.OK);
    }

  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    ui.alert('Error', 
      `An error occurred: ${error.message}\n\n` +
      `Progress has been saved. Run again to resume.`, 
      ui.ButtonSet.OK);
  }
}

/**
 * Shows dialog when there's existing progress
 * Returns: startIndex to begin from, or -1 if cancelled
 */
function showResumeDialog(ui, progress, sections, docId) {
  const totalSections = sections.length;
  const completed = progress.lastProcessedIndex + 1;
  const remaining = totalSections - completed;
  const nextSection = sections[completed];
  const nextLabel = nextSection ? (nextSection.heading || `Section ${completed + 1}`) : `Section ${completed + 1}`;
  
  // Build section preview for "jump to" option
  let sectionPreview = '';
  const previewCount = Math.min(15, totalSections);
  for (let i = 0; i < previewCount; i++) {
    const marker = (i < completed) ? '✓' : (i === completed) ? '▶' : '○';
    const heading = sections[i].heading || '(untitled)';
    const truncated = heading.length > 35 ? heading.substring(0, 35) + '...' : heading;
    sectionPreview += `${marker} ${i + 1}. ${truncated}\n`;
  }
  if (totalSections > 15) {
    sectionPreview += `   ... and ${totalSections - 15} more\n`;
  }

  const result = ui.prompt('Resume or Start Over?',
    `📊 PROGRESS FOUND\n\n` +
    `✓ Completed: ${completed} of ${totalSections} sections\n` +
    `✓ Generated: ${progress.success} infographics\n` +
    `○ Remaining: ${remaining} sections\n\n` +
    `Next up: "${nextLabel}"\n\n` +
    `─────────────────────────────\n` +
    `${sectionPreview}\n` +
    `─────────────────────────────\n\n` +
    `Enter your choice:\n` +
    `• Press OK with empty = CONTINUE from section ${completed + 1}\n` +
    `• Enter a number (1-${totalSections}) = JUMP to that section\n` +
    `• Enter 0 = START OVER from beginning\n` +
    `• Press Cancel = Exit`,
    ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() !== ui.Button.OK) {
    return -1; // Cancelled
  }

  const input = result.getResponseText().trim();
  
  if (input === '') {
    // Continue from where we left off
    return completed;
  }
  
  const num = parseInt(input, 10);
  
  if (isNaN(num)) {
    ui.alert('Invalid Input', 'Please enter a number or leave empty to continue.', ui.ButtonSet.OK);
    return -1;
  }
  
  if (num === 0) {
    // Start over
    clearInfographicProgress(docId);
    return 0;
  }
  
  if (num < 1 || num > totalSections) {
    ui.alert('Invalid Section', `Please enter a number between 1 and ${totalSections}.`, ui.ButtonSet.OK);
    return -1;
  }
  
  // Jump to specific section - clear old progress and set new
  clearInfographicProgress(docId);
  if (num > 1) {
    // Set progress to skip sections before the chosen one
    saveInfographicProgress(docId, num - 2, 0, 0, num - 1, 0, totalSections);
  }
  
  return num - 1;
}

/**
 * Shows dialog when starting fresh (no existing progress)
 * Returns: startIndex to begin from, or -1 if cancelled
 */
function showStartDialog(ui, sections, docId) {
  const totalSections = sections.length;
  
  // Build section preview
  let sectionPreview = '';
  const previewCount = Math.min(15, totalSections);
  for (let i = 0; i < previewCount; i++) {
    const heading = sections[i].heading || '(untitled)';
    const truncated = heading.length > 35 ? heading.substring(0, 35) + '...' : heading;
    sectionPreview += `${i + 1}. ${truncated}\n`;
  }
  if (totalSections > 15) {
    sectionPreview += `... and ${totalSections - 15} more\n`;
  }

  const result = ui.prompt('Generate Infographics',
    `📄 DOCUMENT SECTIONS (${totalSections} total)\n\n` +
    `${sectionPreview}\n` +
    `─────────────────────────────\n\n` +
    `Enter your choice:\n` +
    `• Press OK with empty = START from section 1\n` +
    `• Enter a number (1-${totalSections}) = START from that section\n` +
    `• Press Cancel = Exit`,
    ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() !== ui.Button.OK) {
    return -1; // Cancelled
  }

  const input = result.getResponseText().trim();
  
  if (input === '') {
    return 0; // Start from beginning
  }
  
  const num = parseInt(input, 10);
  
  if (isNaN(num) || num < 1 || num > totalSections) {
    ui.alert('Invalid Section', `Please enter a number between 1 and ${totalSections}.`, ui.ButtonSet.OK);
    return -1;
  }
  
  if (num > 1) {
    // Set progress to skip sections before the chosen one
    saveInfographicProgress(docId, num - 2, 0, 0, num - 1, 0, totalSections);
  }
  
  return num - 1;
}

// ============================================================
// PROGRESS PERSISTENCE FOR RESUME CAPABILITY
// ============================================================

/**
 * Save infographic generation progress
 */
function saveInfographicProgress(docId, lastIndex, success, failed, skipped, totalGenerated, totalSections) {
  const props = PropertiesService.getScriptProperties();
  const progress = {
    docId: docId,
    lastProcessedIndex: lastIndex,
    success: success,
    failed: failed,
    skipped: skipped,
    totalGenerated: totalGenerated || success,
    totalSections: totalSections || 0,
    timestamp: new Date().toISOString()
  };
  props.setProperty('INFOGRAPHIC_PROGRESS_' + docId, JSON.stringify(progress));
  log(`  [Progress saved: section ${lastIndex + 1}/${totalSections || '?'}]`);
}

/**
 * Get saved infographic generation progress
 */
function getInfographicProgress(docId) {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('INFOGRAPHIC_PROGRESS_' + docId);
  if (saved) {
    try {
      return JSON.parse(saved);
    } catch (e) {
      return null;
    }
  }
  return null;
}

/**
 * Clear saved infographic generation progress
 */
function clearInfographicProgress(docId) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('INFOGRAPHIC_PROGRESS_' + docId);
  log('Progress cleared');
}

/**
 * Generate a single document-wide summary infographic
 */
function generateDocumentSummary() {
  const ui = DocumentApp.getUi();
  const doc = DocumentApp.getActiveDocument();

  const apiKey = getApiKey();
  if (!apiKey) {
    ui.alert('API Key Required',
      'Please set your Gemini API key first.\n\nGo to: Infographic Generator → Set API Key',
      ui.ButtonSet.OK);
    return;
  }

  ui.alert('Processing',
    'Generating document summary infographic...',
    ui.ButtonSet.OK);

  try {
    const fullText = getFullDocumentText(doc);
    const title = doc.getName();

    if (fullText.length < CONFIG.MIN_SECTION_LENGTH) {
      ui.alert('No Content', 'Document has insufficient content for summary generation.', ui.ButtonSet.OK);
      return;
    }

    // Generate summary infographic
    const summaryPrompt = createDocumentSummaryPrompt(title, fullText, apiKey);
    const imageBlob = generateInfographicFromPrompt(summaryPrompt, apiKey);

    if (imageBlob) {
      // Insert at beginning of document
      insertImageAtPosition(doc.getBody(), 0, imageBlob, `${title} - Document Overview`);
      ui.alert('Complete', 'Document summary infographic generated successfully!', ui.ButtonSet.OK);
    } else {
      ui.alert('Failed', 'Could not generate document summary infographic.', ui.ButtonSet.OK);
    }

  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    ui.alert('Error', `An error occurred: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Legacy function - redirects to smart analysis
 */
function generateInfographics() {
  generateSmartInfographics();
}

/**
 * Preview analysis without generating images (uses holistic analysis)
 */
function analyzeDocumentPreview() {
  const ui = DocumentApp.getUi();
  const doc = DocumentApp.getActiveDocument();

  const apiKey = getApiKey();
  if (!apiKey) {
    ui.alert('API Key Required',
      'Please set your Gemini API key first.',
      ui.ButtonSet.OK);
    return;
  }

  try {
    log('Starting document analysis preview...');

    const sections = extractSections(doc);
    log(`Extracted ${sections.length} sections`);

    const fullText = getFullDocumentText(doc);
    log(`Document text length: ${fullText.length} characters`);

    const plan = analyzeDocumentHolistically(sections, fullText, apiKey);
    log(`Analysis complete. Infographics planned: ${plan.infographics ? plan.infographics.length : 0}`);

    let message = `Document Analysis Results\n\n`;
    message += `Total sections found: ${sections.length}\n`;
    message += `Recommended infographics: ${plan.infographics.length}\n\n`;

    if (plan.reasoning) {
      message += `Strategy: ${plan.reasoning}\n\n`;
    }

    if (plan.infographics.length > 0) {
      message += `Planned infographics:\n`;
      plan.infographics.forEach((info, i) => {
        const scopeLabel = info.scope === 'document-wide' ? '📄 Document-wide'
          : info.scope === 'multi-section' ? '📑 Multi-section'
          : '📝 Single-section';

        message += `\n${i + 1}. ${scopeLabel}: ${info.title}\n`;
        message += `   Type: ${info.infographicType}\n`;
        message += `   Reason: ${info.reason}\n`;

        if (info.scope === 'multi-section' && Array.isArray(info.sectionIndices)) {
          const sectionNames = info.sectionIndices
            .filter(idx => idx >= 0 && idx < sections.length)
            .map(idx => sections[idx].heading || `Section ${idx + 1}`)
            .join(', ');
          message += `   Spans: ${sectionNames}\n`;
        }
      });
    } else if (plan.reasoning === 'Analysis failed') {
      message += `\nNote: The AI analysis failed to return valid results.\nCheck Apps Script logs (View → Logs) for details.`;
    }

    ui.alert('Analysis Preview', message, ui.ButtonSet.OK);

  } catch (error) {
    log(`Error in analyzeDocumentPreview: ${error.message}`);
    log(`Stack: ${error.stack}`);
    ui.alert('Error', `Analysis failed: ${error.message}\n\nCheck Apps Script logs for details.`, ui.ButtonSet.OK);
  }
}

/**
 * Extracts logical sections from the document
 */
function extractSections(doc) {
  const body = doc.getBody();
  const sections = [];
  let currentSection = {
    text: '',
    startElement: null,
    heading: ''
  };

  const numChildren = body.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    const elementType = element.getType();

    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      const heading = paragraph.getHeading();
      const text = paragraph.getText().trim();

      // Check if this is a heading (new section)
      if (heading !== DocumentApp.ParagraphHeading.NORMAL && text.length > 0) {
        // Save previous section if it has content
        if (currentSection.text.length >= CONFIG.MIN_SECTION_LENGTH) {
          sections.push({...currentSection});
        }

        // Start new section
        currentSection = {
          text: '',
          startElement: element,
          startIndex: i,
          heading: text
        };
      } else if (text.length > 0) {
        // Add to current section
        if (!currentSection.startElement) {
          currentSection.startElement = element;
          currentSection.startIndex = i;
        }
        currentSection.text += (currentSection.text ? '\n' : '') + text;
      }
    } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = element.asListItem();
      const text = listItem.getText().trim();

      if (text.length > 0) {
        if (!currentSection.startElement) {
          currentSection.startElement = element;
          currentSection.startIndex = i;
        }
        currentSection.text += (currentSection.text ? '\n• ' : '• ') + text;
      }
    } else if (elementType === DocumentApp.ElementType.TABLE) {
      // Tables might benefit from infographic representation
      const table = element.asTable();
      let tableText = '[TABLE DATA]\n';

      for (let row = 0; row < table.getNumRows(); row++) {
        const tableRow = table.getRow(row);
        const rowData = [];
        for (let cell = 0; cell < tableRow.getNumCells(); cell++) {
          rowData.push(tableRow.getCell(cell).getText().trim());
        }
        tableText += rowData.join(' | ') + '\n';
      }

      if (!currentSection.startElement) {
        currentSection.startElement = element;
        currentSection.startIndex = i;
      }
      currentSection.text += (currentSection.text ? '\n' : '') + tableText;
    }
  }

  // Don't forget the last section
  if (currentSection.text.length >= CONFIG.MIN_SECTION_LENGTH) {
    sections.push(currentSection);
  }

  return sections;
}

/**
 * Uses Gemini to analyze which sections would benefit from infographics
 * (Legacy function - used by batch mode)
 */
function analyzeSectionsForVisuals(sections, apiKey) {
  const sectionsToVisualize = [];

  // Limit sections to process (use MAX_TOTAL_INFOGRAPHICS as upper bound)
  const sectionsToAnalyze = sections.slice(0, CONFIG.MAX_TOTAL_INFOGRAPHICS);

  for (const section of sectionsToAnalyze) {
    const analysis = analyzeSection(section, apiKey);

    if (analysis.needsInfographic) {
      sectionsToVisualize.push({
        ...section,
        reason: analysis.reason,
        infographicType: analysis.infographicType,
        visualPrompt: analysis.visualPrompt
      });
    }
  }

  return sectionsToVisualize;
}

/**
 * Analyzes a single section using Gemini
 */
function analyzeSection(section, apiKey) {
  const prompt = `Analyze this document section and determine if it would benefit from an infographic visualization.

SECTION CONTENT:
${section.heading ? `Heading: ${section.heading}\n` : ''}
${section.text}

Respond in JSON format:
{
  "needsInfographic": true/false,
  "reason": "Brief explanation of why this section would/wouldn't benefit from a visual",
  "infographicType": "flowchart|comparison|statistics|process|timeline|hierarchy|concept-map|none",
  "visualPrompt": "If needsInfographic is true, provide a detailed prompt for generating an infographic that represents this content. Include specific data points, labels, colors (professional palette), and layout suggestions. Make it suitable for a business/professional document."
}

Consider these factors:
- Statistical data or numbers that could be visualized
- Processes or workflows that could be shown as flowcharts
- Comparisons between items
- Hierarchical relationships
- Timeline or sequential information
- Complex concepts that need visual explanation

Only recommend infographics for content that would genuinely benefit from visual representation.`;

  const response = callGeminiText(prompt, apiKey);

  try {
    const parsed = extractAndParseJSON(response);
    if (parsed) {
      return parsed;
    }
  } catch (e) {
    Logger.log(`Failed to parse analysis response: ${e.message}`);
  }

  return { needsInfographic: false, reason: 'Analysis failed' };
}

/**
 * Gets full document text for holistic analysis
 */
function getFullDocumentText(doc) {
  return doc.getBody().getText();
}

/**
 * Analyzes the entire document to determine optimal infographic strategy
 */
function analyzeDocumentHolistically(sections, fullText, apiKey) {
  const sectionSummaries = sections.map((s, i) => ({
    index: i,
    heading: s.heading || `Section ${i + 1}`,
    preview: s.text.substring(0, 200)
  }));

  const prompt = `Analyze this document and recommend infographics. Consider THREE types:

1. DOCUMENT-WIDE: A single infographic summarizing the entire document (executive summary, overview diagram)
2. MULTI-SECTION: Infographics that span 2+ related sections (e.g., a process that spans "Planning" and "Execution" sections)
3. SINGLE-SECTION: Infographics for individual sections with standalone visual value

DOCUMENT CONTENT:
${fullText.substring(0, 8000)}

SECTION STRUCTURE:
${JSON.stringify(sectionSummaries, null, 2)}

Respond in JSON format:
{
  "infographics": [
    {
      "scope": "document-wide|multi-section|single-section",
      "sectionIndices": [0] or [0,1,2] or "all",
      "title": "Infographic title",
      "reason": "Why this infographic adds value",
      "infographicType": "flowchart|comparison|statistics|process|timeline|hierarchy|concept-map|overview",
      "visualPrompt": "Detailed prompt for generating this infographic, including specific content, data points, labels, and design guidance"
    }
  ],
  "reasoning": "Brief explanation of overall strategy"
}

Guidelines:
- Recommend document-wide infographic ONLY if the document tells a cohesive story or has a central theme
- Group related sections into multi-section infographics when they form a logical unit (e.g., steps in a process)
- Use single-section for standalone data/concepts that don't relate to other sections
- Avoid redundancy - don't create overlapping infographics
- Maximum ${CONFIG.MAX_TOTAL_INFOGRAPHICS} infographics total
- Be selective - only recommend infographics that genuinely add visual value`;

  log('Calling Gemini API for holistic analysis...');
  const response = callGeminiText(prompt, apiKey);

  if (!response) {
    log('ERROR: No response received from Gemini API');
    return { infographics: [], reasoning: 'No API response received' };
  }

  log(`Gemini response length: ${response.length} characters`);
  log(`Response preview: ${response.substring(0, 300)}...`);

  try {
    // Try to extract and parse JSON from the response
    const parsed = extractAndParseJSON(response);
    if (parsed && parsed.infographics) {
      log(`Successfully parsed ${parsed.infographics.length} infographics`);
      return parsed;
    } else {
      log('Parsed result missing infographics array');
      log(`Parsed object keys: ${parsed ? Object.keys(parsed).join(', ') : 'null'}`);
    }
  } catch (e) {
    log(`Failed to parse holistic analysis: ${e.message}`);
  }

  // Log full response for debugging
  log(`Full response for debugging:\n${response}`);

  return { infographics: [], reasoning: 'Analysis failed - check logs for raw response' };
}

/**
 * Robustly extracts and parses JSON from a text response
 */
function extractAndParseJSON(text) {
  if (!text) return null;

  // First, try to find JSON block in markdown code fence
  const codeBlockMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (codeBlockMatch) {
    try {
      return JSON.parse(codeBlockMatch[1].trim());
    } catch (e) {
      // Continue to other methods
    }
  }

  // Try to find the outermost { } pair with balanced braces
  let braceCount = 0;
  let startIndex = -1;
  let endIndex = -1;

  for (let i = 0; i < text.length; i++) {
    if (text[i] === '{') {
      if (braceCount === 0) {
        startIndex = i;
      }
      braceCount++;
    } else if (text[i] === '}') {
      braceCount--;
      if (braceCount === 0 && startIndex !== -1) {
        endIndex = i;
        break;
      }
    }
  }

  if (startIndex !== -1 && endIndex !== -1) {
    let jsonStr = text.substring(startIndex, endIndex + 1);

    // Clean up common issues
    jsonStr = cleanJSONString(jsonStr);

    try {
      return JSON.parse(jsonStr);
    } catch (e) {
      Logger.log(`JSON parse error after cleanup: ${e.message}`);

      // Try a more aggressive cleanup
      try {
        jsonStr = aggressiveJSONCleanup(jsonStr);
        return JSON.parse(jsonStr);
      } catch (e2) {
        Logger.log(`Aggressive cleanup also failed: ${e2.message}`);
      }
    }
  }

  return null;
}

/**
 * Cleans common JSON formatting issues
 */
function cleanJSONString(jsonStr) {
  // Remove trailing commas before ] or }
  jsonStr = jsonStr.replace(/,(\s*[}\]])/g, '$1');

  // Fix unescaped newlines in strings (replace with \n)
  // This is tricky - we need to be careful not to break valid JSON

  // Remove any control characters except valid whitespace
  jsonStr = jsonStr.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

  return jsonStr;
}

/**
 * More aggressive JSON cleanup for stubborn cases
 */
function aggressiveJSONCleanup(jsonStr) {
  // Replace smart quotes with regular quotes
  jsonStr = jsonStr.replace(/[""]/g, '"');
  jsonStr = jsonStr.replace(/['']/g, "'");

  // Remove any text after the last }
  const lastBrace = jsonStr.lastIndexOf('}');
  if (lastBrace !== -1) {
    jsonStr = jsonStr.substring(0, lastBrace + 1);
  }

  // Try to fix arrays with missing commas between objects
  jsonStr = jsonStr.replace(/\}\s*\{/g, '},{');

  // Remove trailing commas (again, after other fixes)
  jsonStr = jsonStr.replace(/,(\s*[}\]])/g, '$1');

  return jsonStr;
}

/**
 * Creates a prompt for document summary infographic
 */
function createDocumentSummaryPrompt(title, fullText, apiKey) {
  // First, ask Gemini to create the visual prompt
  const analysisPrompt = `Create a detailed infographic prompt for this document.

DOCUMENT TITLE: ${title}

DOCUMENT CONTENT:
${fullText.substring(0, 6000)}

Create a detailed prompt for generating an executive summary infographic that captures:
- The main theme or purpose
- Key points or takeaways (3-5 maximum)
- Any important data, statistics, or relationships
- The logical flow or structure

Respond with ONLY the image generation prompt, no other text. Make it detailed and specific for creating a professional business infographic.`;

  return callGeminiText(analysisPrompt, apiKey);
}

/**
 * Generates an infographic from a pre-made prompt
 */
function generateInfographicFromPrompt(visualPrompt, apiKey) {
  const fullPrompt = `${INFOGRAPHIC_STYLE}

CONTENT:
${visualPrompt}`;

  const referenceImages = getReferenceImages();
  return callGeminiImageGeneration(fullPrompt, apiKey, referenceImages);
}

/**
 * Processes the holistic infographic plan (legacy - batch mode)
 */
function processInfographicPlan(doc, sections, plan, apiKey) {
  const body = doc.getBody();
  const results = {
    success: 0,
    failed: 0,
    documentWide: 0,
    multiSection: 0,
    singleSection: 0
  };

  // Sort infographics by insertion position (descending) to maintain correct positions
  const sortedInfographics = plan.infographics.slice().sort((a, b) => {
    const posA = getInsertionPosition(a, sections);
    const posB = getInsertionPosition(b, sections);
    return posB - posA;
  });

  for (const infographic of sortedInfographics) {
    try {
      // Build content for the infographic
      const content = buildInfographicContent(infographic, sections, doc);

      const fullPrompt = `${INFOGRAPHIC_STYLE}

TITLE: ${infographic.title}
INFOGRAPHIC TYPE: ${infographic.infographicType}

CONTENT TO VISUALIZE:
${content}

ADDITIONAL GUIDANCE:
${infographic.visualPrompt}`;

      const referenceImages = getReferenceImages();
      const imageBlob = callGeminiImageGeneration(fullPrompt, apiKey, referenceImages);

      if (imageBlob) {
        const insertPos = getInsertionPosition(infographic, sections);
        insertImageAtPosition(body, insertPos, imageBlob, infographic.title);

        results.success++;
        if (infographic.scope === 'document-wide') results.documentWide++;
        else if (infographic.scope === 'multi-section') results.multiSection++;
        else results.singleSection++;

        Utilities.sleep(2000);
      } else {
        results.failed++;
      }
    } catch (error) {
      Logger.log(`Failed to process infographic: ${error.message}`);
      results.failed++;
    }
  }

  return results;
}

/**
 * Processes the holistic infographic plan INCREMENTALLY with resume capability
 */
function processInfographicPlanIncremental(docId, sections, plan, apiKey, startIndex) {
  const results = {
    success: 0,
    failed: 0,
    documentWide: 0,
    multiSection: 0,
    singleSection: 0,
    completed: false
  };

  // Load any existing progress
  const existingProgress = getSmartProgress(docId);
  if (existingProgress && startIndex > 0) {
    results.success = existingProgress.success || 0;
    results.failed = existingProgress.failed || 0;
    results.documentWide = existingProgress.documentWide || 0;
    results.multiSection = existingProgress.multiSection || 0;
    results.singleSection = existingProgress.singleSection || 0;
  }

  // Sort infographics by insertion position (descending) to maintain correct positions
  const sortedInfographics = plan.infographics.slice().sort((a, b) => {
    const posA = getInsertionPosition(a, sections);
    const posB = getInsertionPosition(b, sections);
    return posB - posA;
  });

  log(`Processing ${sortedInfographics.length} planned infographics, starting from index ${startIndex}`);

  for (let i = startIndex; i < sortedInfographics.length; i++) {
    const infographic = sortedInfographics[i];
    const infoLabel = infographic.title || `Infographic ${i + 1}`;
    
    log(`\n[${i + 1}/${sortedInfographics.length}] Generating: ${infoLabel}`);
    log(`  Type: ${infographic.scope} - ${infographic.infographicType}`);
    
    try {
      // Re-open document to get fresh state
      let doc = DocumentApp.openById(docId);
      const currentSections = extractSections(doc);
      
      // Build content for the infographic
      const content = buildInfographicContent(infographic, currentSections, doc);

      const fullPrompt = `${INFOGRAPHIC_STYLE}

TITLE: ${infographic.title}
INFOGRAPHIC TYPE: ${infographic.infographicType}

CONTENT TO VISUALIZE:
${content}

ADDITIONAL GUIDANCE:
${infographic.visualPrompt}`;

      log(`  Calling image generation API...`);
      const referenceImages = getReferenceImages();
      const imageBlob = callGeminiImageGeneration(fullPrompt, apiKey, referenceImages);

      if (imageBlob) {
        log(`  Inserting image into document...`);
        const body = doc.getBody();
        const insertPos = getInsertionPosition(infographic, currentSections);
        insertImageAtPosition(body, insertPos, imageBlob, infographic.title);
        doc.saveAndClose();

        results.success++;
        if (infographic.scope === 'document-wide') results.documentWide++;
        else if (infographic.scope === 'multi-section') results.multiSection++;
        else results.singleSection++;

        log(`  ✓ Success! (${results.success} total)`);
        
        // Save progress after each successful generation
        saveSmartProgress(docId, i, plan, results);

        Utilities.sleep(2000);
      } else {
        results.failed++;
        log(`  ✗ Failed to generate image`);
        saveSmartProgress(docId, i, plan, results);
      }
    } catch (error) {
      Logger.log(`  ✗ Error: ${error.message}`);
      results.failed++;
      saveSmartProgress(docId, i, plan, results);
      
      // Continue to next infographic rather than aborting
      Utilities.sleep(1000);
    }
  }

  results.completed = true;
  return results;
}

// ============================================================
// SMART ANALYSIS PROGRESS PERSISTENCE
// ============================================================

/**
 * Save smart analysis progress
 */
function saveSmartProgress(docId, lastIndex, plan, results) {
  const props = PropertiesService.getScriptProperties();
  const progress = {
    docId: docId,
    lastProcessedIndex: lastIndex,
    plan: plan,
    success: results.success,
    failed: results.failed,
    documentWide: results.documentWide,
    multiSection: results.multiSection,
    singleSection: results.singleSection,
    timestamp: new Date().toISOString()
  };
  props.setProperty('SMART_PROGRESS_' + docId, JSON.stringify(progress));
  log(`  [Smart progress saved: infographic ${lastIndex + 1}]`);
}

/**
 * Get saved smart analysis progress
 */
function getSmartProgress(docId) {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('SMART_PROGRESS_' + docId);
  if (saved) {
    try {
      return JSON.parse(saved);
    } catch (e) {
      return null;
    }
  }
  return null;
}

/**
 * Clear saved smart analysis progress
 */
function clearSmartProgress(docId) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SMART_PROGRESS_' + docId);
  log('Smart progress cleared');
}


/**
 * Gets the content for an infographic based on its scope
 */
function buildInfographicContent(infographic, sections, doc) {
  if (infographic.sectionIndices === 'all' || infographic.scope === 'document-wide') {
    return getFullDocumentText(doc).substring(0, 5000);
  }

  const indices = Array.isArray(infographic.sectionIndices)
    ? infographic.sectionIndices
    : [infographic.sectionIndices];

  return indices
    .filter(i => i >= 0 && i < sections.length)
    .map(i => {
      const s = sections[i];
      return `${s.heading ? `## ${s.heading}\n` : ''}${s.text}`;
    })
    .join('\n\n');
}

/**
 * Determines where to insert an infographic in the document
 */
function getInsertionPosition(infographic, sections) {
  if (infographic.sectionIndices === 'all' || infographic.scope === 'document-wide') {
    return 0; // Beginning of document
  }

  const indices = Array.isArray(infographic.sectionIndices)
    ? infographic.sectionIndices
    : [infographic.sectionIndices];

  const firstIndex = Math.min(...indices.filter(i => i >= 0 && i < sections.length));

  if (firstIndex >= 0 && firstIndex < sections.length) {
    return sections[firstIndex].startIndex || 0;
  }

  return 0;
}

/**
 * Inserts an image at a specific position with caption
 */
function insertImageAtPosition(body, position, imageBlob, caption) {
  const imageParagraph = body.insertParagraph(position, '');
  imageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const image = imageParagraph.appendInlineImage(imageBlob);

  // Scale image to full page width (468 points = 6.5 inches standard content width)
  const fullPageWidth = 468;
  const originalWidth = image.getWidth();
  const originalHeight = image.getHeight();
  const scale = fullPageWidth / originalWidth;

  image.setWidth(fullPageWidth);
  image.setHeight(Math.round(originalHeight * scale));

  // Add caption
  const captionParagraph = body.insertParagraph(position + 1, `Figure: ${caption}`);
  captionParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  captionParagraph.setItalic(true);
  captionParagraph.setFontSize(10);
  captionParagraph.setForegroundColor(COLORS.TEXT_SECONDARY);

  // Add spacing
  const spacer = body.insertParagraph(position + 2, '');
  spacer.setSpacingAfter(12);
}

/**
 * Processes sections and generates/embeds infographics
 */
function processInfographics(doc, sections, apiKey) {
  const body = doc.getBody();
  let success = 0;
  let failed = 0;

  // Sort sections by startIndex in descending order to maintain positions when inserting
  sections.sort((a, b) => b.startIndex - a.startIndex);

  for (const section of sections) {
    try {
      // Generate infographic
      const imageBlob = generateInfographic(section, apiKey);

      if (imageBlob) {
        // Insert image above the section
        insertImageAboveSection(body, section, imageBlob);
        success++;

        // Small delay to avoid rate limiting
        Utilities.sleep(2000);
      } else {
        failed++;
      }
    } catch (error) {
      Logger.log(`Failed to process section: ${error.message}`);
      failed++;
    }
  }

  return { success, failed };
}

/**
 * Generates an infographic using Gemini Nano Banana Pro
 */
function generateInfographic(section, apiKey) {
  const prompt = `${INFOGRAPHIC_STYLE}

TITLE: ${section.heading || 'Visual Summary'}
INFOGRAPHIC TYPE: ${section.infographicType}

CONTENT TO VISUALIZE:
${section.text}

ADDITIONAL GUIDANCE:
${section.visualPrompt}`;

  const referenceImages = getReferenceImages();
  return callGeminiImageGeneration(prompt, apiKey, referenceImages);
}

/**
 * Calls Gemini API for text generation
 */
function callGeminiText(prompt, apiKey) {
  const url = `${CONFIG.API_BASE_URL}/${CONFIG.TEXT_MODEL}:generateContent?key=${apiKey}`;

  log(`Calling Gemini text API: ${CONFIG.TEXT_MODEL}`);

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: 2048
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    log(`API response code: ${responseCode}`);

    if (responseCode !== 200) {
      log(`API error response: ${responseText}`);
      throw new Error(`API returned status ${responseCode}: ${responseText.substring(0, 200)}`);
    }

    const result = JSON.parse(responseText);

    if (result.error) {
      log(`Gemini API error: ${JSON.stringify(result.error)}`);
      throw new Error(`Gemini API error: ${result.error.message}`);
    }

    if (!result.candidates || result.candidates.length === 0) {
      log(`No candidates in response: ${JSON.stringify(result)}`);
      throw new Error('No candidates returned from Gemini API');
    }

    const text = result.candidates[0].content.parts[0].text;
    log(`Successfully received ${text.length} characters from API`);
    return text;

  } catch (e) {
    log(`Exception in callGeminiText: ${e.message}`);
    throw e;
  }
}

/**
 * Calls Gemini API for image generation (Nano Banana Pro)
 * Supports optional reference images
 */
function callGeminiImageGeneration(prompt, apiKey, referenceImages) {
  const url = `${CONFIG.API_BASE_URL}/${CONFIG.IMAGE_MODEL}:generateContent?key=${apiKey}`;

  // Build parts array with text prompt and optional reference images
  const parts = [];

  // Add reference images first (if any)
  if (referenceImages && referenceImages.length > 0) {
    for (const refImage of referenceImages) {
      parts.push({
        inlineData: {
          mimeType: refImage.mimeType,
          data: refImage.base64Data
        }
      });
    }
  }

  // Add the text prompt
  parts.push({ text: prompt });

  const payload = {
    contents: [{
      parts: parts
    }],
    generationConfig: {
      responseModalities: ['image', 'text'],
      imageConfig: {
        aspectRatio: '16:9'
      }
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    Logger.log(`Image generation error: ${result.error.message}`);
    return null;
  }

  // Extract image data from response
  const responseParts = result.candidates[0].content.parts;
  for (const part of responseParts) {
    if (part.inlineData && part.inlineData.mimeType.startsWith('image/')) {
      const imageBytes = Utilities.base64Decode(part.inlineData.data);
      return Utilities.newBlob(imageBytes, part.inlineData.mimeType, 'infographic.png');
    }
  }

  Logger.log('No image found in response');
  return null;
}

/**
 * Fetches an image from URL and returns it as base64 with mimeType
 */
function fetchImageAsBase64(imageUrl) {
  try {
    const response = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`Failed to fetch image from ${imageUrl}: HTTP ${responseCode}`);
      return null;
    }

    const blob = response.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();

    return {
      base64Data: base64Data,
      mimeType: mimeType
    };
  } catch (error) {
    Logger.log(`Error fetching image: ${error.message}`);
    return null;
  }
}

/**
 * Gets reference images for infographic generation (e.g., logo)
 * Caches the logo to avoid repeated fetches
 */
let cachedLogo = null;
function getReferenceImages() {
  if (!BRANDING.INCLUDE_LOGO) {
    return [];
  }

  // Use cached logo if available
  if (cachedLogo) {
    return [cachedLogo];
  }

  const logo = fetchImageAsBase64(BRANDING.LOGO_URL);
  if (logo) {
    cachedLogo = logo;
    return [logo];
  }

  return [];
}

/**
 * Inserts an image above a section in the document
 */
function insertImageAboveSection(body, section, imageBlob) {
  const insertIndex = section.startIndex;

  // Insert a new paragraph for the image
  const imageParagraph = body.insertParagraph(insertIndex, '');
  imageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // Insert the image
  const image = imageParagraph.appendInlineImage(imageBlob);

  // Scale image to full page width (468 points = 6.5 inches standard content width)
  const fullPageWidth = 468;
  const originalWidth = image.getWidth();
  const originalHeight = image.getHeight();
  const scale = fullPageWidth / originalWidth;

  image.setWidth(fullPageWidth);
  image.setHeight(Math.round(originalHeight * scale));

  // Add a caption
  const captionText = section.heading
    ? `Figure: ${section.heading} - Visual Summary`
    : 'Figure: Visual Summary';

  const caption = body.insertParagraph(insertIndex + 1, captionText);
  caption.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  caption.setItalic(true);
  caption.setFontSize(10);
  caption.setForegroundColor(COLORS.TEXT_SECONDARY);

  // Add spacing after caption
  const spacer = body.insertParagraph(insertIndex + 2, '');
  spacer.setSpacingAfter(12);
}

/**
 * Gets API key from script properties
 */
function getApiKey() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty(CONFIG.GEMINI_API_KEY_PROPERTY);
}

/**
 * Shows dialog to set API key
 */
function showApiKeyDialog() {
  const ui = DocumentApp.getUi();
  const result = ui.prompt(
    'Set Gemini API Key',
    'Enter your Gemini API key (get one at https://aistudio.google.com/apikey):',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const apiKey = result.getResponseText().trim();
    if (apiKey) {
      PropertiesService.getScriptProperties().setProperty(
        CONFIG.GEMINI_API_KEY_PROPERTY,
        apiKey
      );
      ui.alert('Success', 'API key saved successfully!', ui.ButtonSet.OK);
    }
  }
}

/**
 * Shows help information for Book Formatting
 */
function showFormattingHelp() {
  const ui = DocumentApp.getUi();
  const helpText = `
BOOK FORMATTING HELP

Formats your document with professional styling.

FORMAT MODES:
• Fresh - First-time formatting, applies base
  font globally then formats all elements
• Resume - Skips already-formatted elements,
  use if interrupted mid-process

HEADING STYLES:
• H1: ${FONTS.PRIMARY} ${STYLES.sizes.h1}pt, red bar prefix
• H2: ${FONTS.PRIMARY} ${STYLES.sizes.h2}pt, arrow prefix
• H3: ${FONTS.PRIMARY} ${STYLES.sizes.h3}pt, bold
• H4: ${FONTS.SECONDARY} ${STYLES.sizes.h4}pt, normal
• H5: ${FONTS.PRIMARY} ${STYLES.sizes.h5}pt, bold

FEATURES:
• Auto-bolds list terms before : or -
• Tables get header styling & zebra stripes
• Links restored to blue with underline
• Images resized to full width

UTILITIES:
• Promote Headings - Moves all up one level
• Remove Horizontal Lines - Cleans up dividers
• Remove HTML Comments - Strips <!-- -->
• Remove Bookmarks - Clears all bookmarks

TIPS:
• Use Preview (First 5) to test before full run
• Individual Formatters let you redo specific
  element types without re-running everything
  `;

  ui.alert('Book Formatting Help', helpText, ui.ButtonSet.OK);
}

/**
 * Shows help information for Infographic Generator
 */
function showInfographicHelp() {
  const ui = DocumentApp.getUi();
  const helpText = `
INFOGRAPHIC GENERATOR HELP

Analyzes your document and generates infographics
using Google Gemini's image generation model.

HOW TO USE:

1. Click "▶️ Generate Infographics"

2. You'll see a list of all sections in your doc

3. Choose where to start:
   • Leave empty → Start from section 1
   • Enter a number → Jump to that section
   • Enter 0 → Start over (if resuming)

4. Images generate one at a time and appear
   in your document as they complete

⏱️ TIMEOUT HANDLING:

• Progress saves automatically after each image
• If it times out, just run again
• You'll see your progress and can continue
  or jump to any section

INFOGRAPHIC TYPES DETECTED:
• Flowcharts & processes
• Statistics & data
• Comparisons
• Timelines
• Hierarchies
• Concept maps

TIPS:
• Use "Analyze Document" to preview first
• Each run processes ~10 sections
• Large docs may need multiple runs

API KEY:
Get one at: https://aistudio.google.com/apikey
  `;

  ui.alert('Help', helpText, ui.ButtonSet.OK);
}

/**
 * Test function for development
 */
function testApiConnection() {
  const apiKey = getApiKey();
  if (!apiKey) {
    Logger.log('No API key set');
    return;
  }

  try {
    const response = callGeminiText('Say "API connection successful" in exactly those words.', apiKey);
    Logger.log(`API Test Response: ${response}`);
  } catch (error) {
    Logger.log(`API Test Failed: ${error.message}`);
  }
}

// ============================================================
// DOCUMENT FORMATTING - MAIN ENTRY POINTS
// ============================================================

/**
 * Fresh format - fastest for first-time formatting.
 */
function formatAllFresh() {
  formatAll(true);
}

/**
 * Resume format - use if previously interrupted.
 */
function formatAllResume() {
  formatAll(false);
}

/**
 * Main formatting orchestrator.
 */
function formatAll(skipChecks) {
  const ui = DocumentApp.getUi();
  const start = new Date();
  const mode = skipChecks ? 'FRESH' : 'RESUME';

  log(`\n${'='.repeat(50)}`);
  log(`Starting full document format - ${mode} mode`);
  log(`${'='.repeat(50)}`);

  if (skipChecks) {
    applyBaseFontToDocument();
  }

  formatH1Only(skipChecks);
  formatH2Only(skipChecks);
  formatH3Only(skipChecks);
  formatH4Only(skipChecks);
  formatH5Only(skipChecks);

  if (!skipChecks) {
    formatBodyOnly(skipChecks);
  }

  formatListsOnly(skipChecks);
  formatTablesOnly(skipChecks);
  resizeImagesToFullWidth(skipChecks);
  cleanupComments();
  restoreLinkFormatting();

  const elapsed = ((new Date()) - start) / 1000;
  log(`\n${'='.repeat(50)}`);
  log(`Complete! Total time: ${elapsed.toFixed(1)} seconds`);
  log(`${'='.repeat(50)}`);

  ui.alert('Formatting Complete',
    `Document formatted successfully!\n\nMode: ${mode}\nTime: ${elapsed.toFixed(1)} seconds`,
    ui.ButtonSet.OK);
}

// ============================================================
// BASE FONT APPLICATION
// ============================================================

function applyBaseFontToDocument() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  log('\nApplying base font to entire document...');
  body.setFontFamily(STYLES.fonts.body);
  body.setFontSize(STYLES.sizes.body);
  body.setForegroundColor(STYLES.colors.charcoal);
  log('  Base font applied');
}

// ============================================================
// H1 FORMATTING
// ============================================================

function formatH1Only(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let paragraphs = body.getParagraphs();

  let total = 0;
  for (const para of paragraphs) {
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
      const text = para.getText();
      if (text.trim() && !text.startsWith('<!--')) total++;
    }
  }

  log(`\nFormatting H1 headings: ${total} found`);
  if (total === 0) return;

  let count = 0;
  let skipped = 0;
  const BATCH_SIZE = 100;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING1) continue;

    let text = para.getText();
    if (!text.trim() || text.startsWith('<!--')) continue;

    if (!skipChecks && text.match(/^┃[^\s]/) && para.getFontSize() === STYLES.sizes.h1) {
      skipped++;
      continue;
    }

    text = text.replace(/^[▌▐│┃|▏▎]+\s*/g, '').trim();
    const newText = '┃' + text;

    para.setAttributes({});
    para.setText(newText);
    para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    para.setFontFamily(STYLES.fonts.h1);
    para.setFontSize(STYLES.sizes.h1);
    para.setBold(true);
    para.setIndentStart(-8);
    para.setIndentFirstLine(-8);
    para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    para.setSpacingBefore(STYLES.spacing.h1Before);
    para.setSpacingAfter(STYLES.spacing.h1After);

    const textEl = para.editAsText();
    textEl.setFontSize(0, 0, STYLES.sizes.h1);
    textEl.setForegroundColor(0, 0, STYLES.colors.headingRed);

    if (newText.length > 1) {
      textEl.setFontSize(1, newText.length - 1, STYLES.sizes.h1);
      textEl.setForegroundColor(1, newText.length - 1, STYLES.colors.headingDark);
    }

    count++;

    if (count % BATCH_SIZE === 0) {
      log(`  Progress: ${count}/${total} formatted`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      paragraphs = body.getParagraphs();
    }
  }

  doc.saveAndClose();
  log(`  H1 complete: ${count} formatted` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

// ============================================================
// H2 FORMATTING
// ============================================================

function formatH2Only(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let paragraphs = body.getParagraphs();

  let total = 0;
  for (const para of paragraphs) {
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING2) {
      const text = para.getText();
      if (text.trim() && !text.startsWith('<!--')) total++;
    }
  }

  log(`\nFormatting H2 headings: ${total} found`);
  if (total === 0) return;

  let count = 0;
  let skipped = 0;
  const BATCH_SIZE = 150;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING2) continue;

    let text = para.getText();
    if (!text.trim() || text.startsWith('<!--')) continue;

    if (!skipChecks && text.match(/^→[^\s]/) && para.getFontSize() === STYLES.sizes.h2) {
      skipped++;
      continue;
    }

    text = text.replace(/^[▌▐│┃|▶▸→➔➜]+\s*/g, '').trim();
    const newText = '→' + text;

    para.setText(newText);
    para.setFontFamily(STYLES.fonts.h2);
    para.setFontSize(STYLES.sizes.h2);
    para.setForegroundColor(STYLES.colors.headingRed);
    para.setBold(true);
    para.setSpacingBefore(STYLES.spacing.h2Before);
    para.setSpacingAfter(STYLES.spacing.h2After);

    count++;

    if (count % BATCH_SIZE === 0) {
      log(`  Progress: ${count}/${total} formatted`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      paragraphs = body.getParagraphs();
    }
  }

  doc.saveAndClose();
  log(`  H2 complete: ${count} formatted` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

// ============================================================
// H3 FORMATTING
// ============================================================

function formatH3Only(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let paragraphs = body.getParagraphs();

  let total = 0;
  for (const para of paragraphs) {
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING3) {
      const text = para.getText();
      if (text.trim() && !text.startsWith('<!--')) total++;
    }
  }

  log(`\nFormatting H3 headings: ${total} found`);
  if (total === 0) return;

  let count = 0;
  let skipped = 0;
  const BATCH_SIZE = 200;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING3) continue;

    const text = para.getText();
    if (!text.trim() || text.startsWith('<!--')) continue;

    if (!skipChecks && para.getFontSize() === STYLES.sizes.h3 && para.getFontFamily() === STYLES.fonts.h3) {
      skipped++;
      continue;
    }

    para.setFontFamily(STYLES.fonts.h3);
    para.setFontSize(STYLES.sizes.h3);
    para.setForegroundColor(STYLES.colors.headingDark);
    para.setBold(true);
    para.setSpacingBefore(STYLES.spacing.h3Before);
    para.setSpacingAfter(STYLES.spacing.h3After);

    count++;

    if (count % BATCH_SIZE === 0) {
      log(`  Progress: ${count}/${total} formatted`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      paragraphs = body.getParagraphs();
    }
  }

  doc.saveAndClose();
  log(`  H3 complete: ${count} formatted` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

// ============================================================
// H4 FORMATTING
// ============================================================

function formatH4Only(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let paragraphs = body.getParagraphs();

  let total = 0;
  for (const para of paragraphs) {
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING4) {
      const text = para.getText();
      if (text.trim() && !text.startsWith('<!--')) total++;
    }
  }

  log(`\nFormatting H4 headings: ${total} found`);
  if (total === 0) return;

  let count = 0;
  let skipped = 0;
  const BATCH_SIZE = 200;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING4) continue;

    const text = para.getText();
    if (!text.trim() || text.startsWith('<!--')) continue;

    if (!skipChecks && para.getFontSize() === STYLES.sizes.h4 && para.getFontFamily() === STYLES.fonts.h4) {
      skipped++;
      continue;
    }

    para.setFontFamily(STYLES.fonts.h4);
    para.setFontSize(STYLES.sizes.h4);
    para.setForegroundColor(STYLES.colors.headingDark);
    para.setBold(false);
    para.setSpacingBefore(STYLES.spacing.h3Before);
    para.setSpacingAfter(STYLES.spacing.h3After);

    count++;

    if (count % BATCH_SIZE === 0) {
      log(`  Progress: ${count}/${total} formatted`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      paragraphs = body.getParagraphs();
    }
  }

  doc.saveAndClose();
  log(`  H4 complete: ${count} formatted` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

// ============================================================
// H5 FORMATTING
// ============================================================

function formatH5Only(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let paragraphs = body.getParagraphs();

  let total = 0;
  for (const para of paragraphs) {
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING5) {
      const text = para.getText();
      if (text.trim() && !text.startsWith('<!--')) total++;
    }
  }

  log(`\nFormatting H5 headings: ${total} found`);
  if (total === 0) return;

  let count = 0;
  let skipped = 0;
  const BATCH_SIZE = 200;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (para.getHeading() !== DocumentApp.ParagraphHeading.HEADING5) continue;

    const text = para.getText();
    if (!text.trim() || text.startsWith('<!--')) continue;

    if (!skipChecks && para.getFontSize() === STYLES.sizes.h5 && para.getFontFamily() === STYLES.fonts.h5) {
      skipped++;
      continue;
    }

    para.setFontFamily(STYLES.fonts.h5);
    para.setFontSize(STYLES.sizes.h5);
    para.setForegroundColor(STYLES.colors.headingDark);
    para.setBold(true);
    para.setSpacingBefore(STYLES.spacing.h3Before);
    para.setSpacingAfter(STYLES.spacing.h3After);

    count++;

    if (count % BATCH_SIZE === 0) {
      log(`  Progress: ${count}/${total} formatted`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      paragraphs = body.getParagraphs();
    }
  }

  doc.saveAndClose();
  log(`  H5 complete: ${count} formatted` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

// ============================================================
// BODY TEXT FORMATTING
// ============================================================

function formatBodyOnly(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let paragraphs = body.getParagraphs();

  let total = 0;
  for (const para of paragraphs) {
    if (para.getHeading() === DocumentApp.ParagraphHeading.NORMAL) {
      const text = para.getText();
      if (!text.trim() || text.startsWith('<!--')) continue;
      const font = para.getFontFamily();
      if (font && (font.toLowerCase().includes('mono') || font.toLowerCase().includes('courier'))) continue;
      if (text.startsWith('    ')) continue;
      total++;
    }
  }

  log(`\nFormatting body paragraphs: ${total} found`);
  if (total === 0) return;

  let count = 0;
  let skipped = 0;
  const BATCH_SIZE = 200;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) continue;

    const text = para.getText();
    if (!text.trim() || text.startsWith('<!--')) continue;

    const font = para.getFontFamily();
    if (font && (font.toLowerCase().includes('mono') || font.toLowerCase().includes('courier'))) continue;
    if (text.startsWith('    ')) continue;

    if (!skipChecks && para.getFontSize() === STYLES.sizes.body && font === STYLES.fonts.body) {
      skipped++;
      continue;
    }

    para.setFontFamily(STYLES.fonts.body);
    para.setFontSize(STYLES.sizes.body);
    para.setBold(false);
    para.setForegroundColor(STYLES.colors.charcoal);
    para.setLineSpacing(STYLES.spacing.lineSpacing);
    para.setSpacingAfter(STYLES.spacing.paragraphAfter);

    count++;

    if (count % BATCH_SIZE === 0) {
      log(`  Progress: ${count}/${total} formatted`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      paragraphs = body.getParagraphs();
    }
  }

  doc.saveAndClose();
  log(`  Body complete: ${count} formatted` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

// ============================================================
// LIST FORMATTING
// ============================================================

function formatListsOnly(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let listItems = body.getListItems();

  const total = listItems.length;
  log(`\nFormatting list items: ${total} found`);
  if (total === 0) return;

  let count = 0;
  let skipped = 0;
  let bolded = 0;
  const BATCH_SIZE = 200;

  for (let i = 0; i < listItems.length; i++) {
    const item = listItems[i];

    if (!skipChecks && item.getSpacingAfter() === 0 && item.getSpacingBefore() === 0) {
      skipped++;
      continue;
    }

    item.setAttributes({});
    item.setFontFamily(STYLES.fonts.body);
    item.setFontSize(STYLES.sizes.body);
    item.setForegroundColor(STYLES.colors.charcoal);
    item.setBold(false);
    item.setItalic(false);
    item.setSpacingBefore(0);
    item.setSpacingAfter(0);
    item.setLineSpacing(1.0);

    const text = item.getText();
    const textElement = item.editAsText();
    const boldEndIndex = findDefinitionDelimiter(text);

    if (boldEndIndex > 0) {
      textElement.setBold(0, boldEndIndex, true);
      bolded++;
    }

    count++;

    if (count % BATCH_SIZE === 0) {
      log(`  Progress: ${count}/${total} formatted`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      listItems = body.getListItems();
    }
  }

  doc.saveAndClose();
  log(`  Lists complete: ${count} formatted, ${bolded} with bold terms` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

function findDefinitionDelimiter(text) {
  const MAX_TERM_LENGTH = 60;

  const colonIndex = text.indexOf(':');
  if (colonIndex > 0 && colonIndex <= MAX_TERM_LENGTH) {
    return colonIndex;
  }

  const dashIndex = text.indexOf(' - ');
  if (dashIndex > 0 && dashIndex <= MAX_TERM_LENGTH) {
    return dashIndex + 2;
  }

  const enDashIndex = text.indexOf(' – ');
  if (enDashIndex > 0 && enDashIndex <= MAX_TERM_LENGTH) {
    return enDashIndex + 2;
  }

  const emDashIndex = text.indexOf(' — ');
  if (emDashIndex > 0 && emDashIndex <= MAX_TERM_LENGTH) {
    return emDashIndex + 2;
  }

  return -1;
}

function boldListTermsOnly() {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let listItems = body.getListItems();

  const total = listItems.length;
  log(`\nBolding list terms: ${total} list items found`);
  if (total === 0) return;

  let bolded = 0;
  const BATCH_SIZE = 200;

  for (let i = 0; i < listItems.length; i++) {
    const item = listItems[i];
    const text = item.getText();
    const textElement = item.editAsText();
    const boldEndIndex = findDefinitionDelimiter(text);

    if (boldEndIndex > 0) {
      textElement.setBold(0, boldEndIndex, true);
      bolded++;
    }

    if ((i + 1) % BATCH_SIZE === 0) {
      log(`  Progress: ${i + 1}/${total} checked`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      listItems = body.getListItems();
    }
  }

  doc.saveAndClose();
  log(`  Complete: ${bolded} list items with bold terms`);

  DocumentApp.getUi().alert('Bold List Terms Complete',
    `Processed ${total} list items.\n${bolded} items had terms bolded.`,
    DocumentApp.getUi().ButtonSet.OK);
}

// ============================================================
// TABLE FORMATTING
// ============================================================

function formatTablesOnly(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  const numTables = body.getTables().length;

  log(`\nFormatting tables: ${numTables} found`);
  if (numTables === 0) return;

  let count = 0;
  let skipped = 0;

  for (let t = 0; t < numTables; t++) {
    const table = body.getTables()[t];

    if (!skipChecks) {
      const firstCell = table.getRow(0).getCell(0);
      if (firstCell.getBackgroundColor() === STYLES.colors.tableHeader) {
        skipped++;
        continue;
      }
    }

    formatSingleTable(table);
    count++;

    if (count % 50 === 0) {
      log(`  Progress: ${count}/${numTables} formatted`);
      doc.saveAndClose();
      doc = DocumentApp.openById(docId);
      body = doc.getBody();
    }
  }

  doc.saveAndClose();
  log(`  Tables complete: ${count} formatted` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

function formatSingleTable(table) {
  const numRows = table.getNumRows();
  const numCols = table.getRow(0).getNumCells();

  table.setBorderWidth(0.25);
  table.setBorderColor(STYLES.colors.tableBorder);

  const fullWidth = 468;
  const minColWidth = 40;

  const colMaxChars = [];
  for (let col = 0; col < numCols; col++) {
    let maxChars = 5;
    for (let row = 0; row < numRows; row++) {
      const cell = table.getRow(row).getCell(col);
      const lines = cell.getText().split('\n');
      for (const line of lines) {
        maxChars = Math.max(maxChars, line.length);
      }
    }
    colMaxChars.push(maxChars);
  }

  const totalChars = colMaxChars.reduce((sum, c) => sum + c, 0);
  for (let col = 0; col < numCols; col++) {
    let colWidth = Math.max(Math.floor((colMaxChars[col] / totalChars) * fullWidth), minColWidth);
    table.setColumnWidth(col, colWidth);
  }

  for (let row = 0; row < numRows; row++) {
    const tableRow = table.getRow(row);
    const isHeader = (row === 0);
    const isOddRow = (row % 2 === 1);

    for (let col = 0; col < tableRow.getNumCells(); col++) {
      const cell = tableRow.getCell(col);

      if (isHeader) {
        cell.setBackgroundColor(STYLES.colors.tableHeader);
      } else if (isOddRow) {
        cell.setBackgroundColor(STYLES.colors.tableAltRow);
      } else {
        cell.setBackgroundColor(STYLES.colors.white);
      }

      cell.setPaddingTop(4);
      cell.setPaddingBottom(4);
      cell.setPaddingLeft(8);
      cell.setPaddingRight(8);

      for (let c = 0; c < cell.getNumChildren(); c++) {
        const child = cell.getChild(c);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const para = child.asParagraph();
          para.setSpacingBefore(0);
          para.setSpacingAfter(0);
          para.setLineSpacing(1.0);
          para.setFontFamily(STYLES.fonts.body);
          para.setFontSize(9);

          if (isHeader) {
            para.setForegroundColor(STYLES.colors.tableHeaderText);
            para.setBold(true);
          } else {
            para.setForegroundColor(STYLES.colors.charcoal);
            para.setBold(false);
          }
        }
      }
    }
  }
}

// ============================================================
// LINK FORMATTING
// ============================================================

function restoreLinkFormatting() {
  const docId = DocumentApp.getActiveDocument().getId();
  const linkColor = STYLES.colors.linkBlue;
  const CHUNK_SIZE = 2000;

  log('\nRestoring link formatting...');

  let totalLinkCount = 0;

  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  const totalParagraphs = body.getParagraphs().length;
  const numParagraphChunks = Math.ceil(totalParagraphs / CHUNK_SIZE);

  log(`  Scanning ${totalParagraphs} paragraphs in ${numParagraphChunks} chunks...`);

  for (let chunk = 0; chunk < numParagraphChunks; chunk++) {
    const startIdx = chunk * CHUNK_SIZE;
    const endIdx = Math.min(startIdx + CHUNK_SIZE, totalParagraphs);
    let chunkLinkCount = 0;

    doc = DocumentApp.openById(docId);
    body = doc.getBody();
    const paragraphs = body.getParagraphs();

    for (let p = startIdx; p < endIdx && p < paragraphs.length; p++) {
      const para = paragraphs[p];
      const text = para.editAsText();
      const textContent = para.getText();

      if (!textContent) continue;

      const indices = text.getTextAttributeIndices();
      for (let j = 0; j < indices.length; j++) {
        const idx = indices[j];
        const url = text.getLinkUrl(idx);
        if (url) {
          const linkEnd = (j + 1 < indices.length) ? indices[j + 1] - 1 : textContent.length - 1;
          text.setForegroundColor(idx, linkEnd, linkColor);
          text.setUnderline(idx, linkEnd, true);
          chunkLinkCount++;
        }
      }
    }

    doc.saveAndClose();
    totalLinkCount += chunkLinkCount;
  }

  doc = DocumentApp.openById(docId);
  body = doc.getBody();
  const totalListItems = body.getListItems().length;

  if (totalListItems > 0) {
    const numListChunks = Math.ceil(totalListItems / CHUNK_SIZE);

    for (let chunk = 0; chunk < numListChunks; chunk++) {
      const startIdx = chunk * CHUNK_SIZE;
      const endIdx = Math.min(startIdx + CHUNK_SIZE, totalListItems);
      let chunkLinkCount = 0;

      doc = DocumentApp.openById(docId);
      body = doc.getBody();
      const listItems = body.getListItems();

      for (let li = startIdx; li < endIdx && li < listItems.length; li++) {
        const item = listItems[li];
        const text = item.editAsText();
        const textContent = item.getText();

        if (!textContent) continue;

        const indices = text.getTextAttributeIndices();
        for (let j = 0; j < indices.length; j++) {
          const idx = indices[j];
          const url = text.getLinkUrl(idx);
          if (url) {
            const linkEnd = (j + 1 < indices.length) ? indices[j + 1] - 1 : textContent.length - 1;
            text.setForegroundColor(idx, linkEnd, linkColor);
            text.setUnderline(idx, linkEnd, true);
            chunkLinkCount++;
          }
        }
      }

      doc.saveAndClose();
      totalLinkCount += chunkLinkCount;
    }
  }

  log(`  Links complete: ${totalLinkCount} restored`);
}

// ============================================================
// IMAGE RESIZING
// ============================================================

function resizeImagesOnly() {
  resizeImagesToFullWidth(true);
}

function resizeImagesToFullWidth(skipChecks) {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  let paragraphs = body.getParagraphs();
  const targetWidth = STYLES.imageWidth;

  let total = 0;
  for (const para of paragraphs) {
    for (let i = 0; i < para.getNumChildren(); i++) {
      const child = para.getChild(i);
      if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
        total++;
      }
    }
  }

  log(`\nResizing images: ${total} found`);
  if (total === 0) return;

  let count = 0;
  let skipped = 0;
  const BATCH_SIZE = 50;

  for (let p = 0; p < paragraphs.length; p++) {
    const para = paragraphs[p];

    for (let i = 0; i < para.getNumChildren(); i++) {
      const child = para.getChild(i);

      if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
        const img = child.asInlineImage();
        const origWidth = img.getWidth();

        if (!skipChecks && origWidth === targetWidth) {
          skipped++;
          continue;
        }

        const origHeight = img.getHeight();
        const aspectRatio = origHeight / origWidth;
        const newHeight = Math.round(targetWidth * aspectRatio);

        img.setWidth(targetWidth);
        img.setHeight(newHeight);
        para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        count++;

        if (count % BATCH_SIZE === 0) {
          log(`  Progress: ${count}/${total} resized`);
          doc.saveAndClose();
          doc = DocumentApp.openById(docId);
          body = doc.getBody();
          paragraphs = body.getParagraphs();
        }
      }
    }
  }

  doc.saveAndClose();
  log(`  Images complete: ${count} resized to ${targetWidth}pt width` + (skipped > 0 ? `, ${skipped} skipped` : ''));
}

// ============================================================
// CLEANUP UTILITIES
// ============================================================

function cleanupComments() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  log('\nRemoving HTML comments...');
  let removed = 0;

  for (let i = paragraphs.length - 1; i >= 0; i--) {
    const text = paragraphs[i].getText().trim();
    if (text.startsWith('<!--') && text.endsWith('-->')) {
      paragraphs[i].removeFromParent();
      removed++;
    }
  }

  log(`  Comments removed: ${removed}`);
}

function removeHorizontalLines() {
  const docId = DocumentApp.getActiveDocument().getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();

  log('\nRemoving horizontal lines...');
  let removed = 0;

  const numChildren = body.getNumChildren();

  // Go backwards to avoid index shifting when removing
  for (let i = numChildren - 1; i >= 0; i--) {
    const child = body.getChild(i);
    const type = child.getType();

    // Check if direct child is a horizontal rule
    if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
      child.removeFromParent();
      removed++;
      continue;
    }

    // Check paragraphs for horizontal rules inside them or text-based lines
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      const text = para.getText().trim();

      // Check for text-based horizontal lines (---, ___, etc.)
      if (/^[-_─━—–=*]{3,}$/.test(text)) {
        para.removeFromParent();
        removed++;
        continue;
      }

      // Check for horizontal rule elements inside the paragraph
      let hasOnlyHorizontalRule = false;
      const numParaChildren = para.getNumChildren();

      for (let j = numParaChildren - 1; j >= 0; j--) {
        try {
          const paraChild = para.getChild(j);
          if (paraChild.getType() === DocumentApp.ElementType.HORIZONTAL_RULE) {
            paraChild.removeFromParent();
            removed++;
            hasOnlyHorizontalRule = true;
          }
        } catch (e) {
          // Some elements may not support getType()
          continue;
        }
      }

      // If paragraph is now empty after removing HR, remove it too
      if (hasOnlyHorizontalRule && para.getText().trim() === '') {
        try {
          para.removeFromParent();
        } catch (e) {
          // Paragraph may already be removed
        }
      }
    }
  }

  // Also search using findElement for any we might have missed
  let searchResult = body.findElement(DocumentApp.ElementType.HORIZONTAL_RULE);
  while (searchResult) {
    try {
      const element = searchResult.getElement();
      const parent = element.getParent();
      element.removeFromParent();
      removed++;

      // Remove empty parent paragraph if applicable
      if (parent && parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
        if (parent.asParagraph().getText().trim() === '') {
          try {
            parent.removeFromParent();
          } catch (e) {}
        }
      }
    } catch (e) {
      // Element may have already been removed
    }
    searchResult = body.findElement(DocumentApp.ElementType.HORIZONTAL_RULE, searchResult);
  }

  doc.saveAndClose();
  log(`  Horizontal lines removed: ${removed}`);

  DocumentApp.getUi().alert('Remove Horizontal Lines',
    `Removed ${removed} horizontal lines from the document.`,
    DocumentApp.getUi().ButtonSet.OK);
}

function removeAllBookmarks() {
  const doc = DocumentApp.getActiveDocument();
  const bookmarks = doc.getBookmarks();

  log('\nRemoving bookmarks...');
  let removed = 0;

  for (const bookmark of bookmarks) {
    bookmark.remove();
    removed++;
  }

  log(`  Bookmarks removed: ${removed}`);
}

// ============================================================
// HEADING PROMOTION
// ============================================================

function promoteHeadings() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Promote Headings',
    'This will promote all headings up by one level:\n\n' +
    '• H2 → H1\n• H3 → H2\n• H4 → H3\n• H5 → H4\n• H1 stays as H1\n\nContinue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Cancelled', 'No changes made.', ui.ButtonSet.OK);
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  let h2Count = 0, h3Count = 0, h4Count = 0, h5Count = 0;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const heading = para.getHeading();

    if (heading === DocumentApp.ParagraphHeading.HEADING5) {
      para.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      h5Count++;
    }
    else if (heading === DocumentApp.ParagraphHeading.HEADING4) {
      para.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      h4Count++;
    }
    else if (heading === DocumentApp.ParagraphHeading.HEADING3) {
      para.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      h3Count++;
    }
    else if (heading === DocumentApp.ParagraphHeading.HEADING2) {
      para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      h2Count++;
    }
  }

  const total = h2Count + h3Count + h4Count + h5Count;
  log(`\nHeadings promoted: ${total} total`);

  ui.alert(
    'Headings Promoted',
    `${total} headings were promoted:\n\n` +
    `• H2 → H1: ${h2Count}\n• H3 → H2: ${h3Count}\n• H4 → H3: ${h4Count}\n• H5 → H4: ${h5Count}\n\n` +
    'Run "Format Everything (Fresh)" to apply styling.',
    ui.ButtonSet.OK
  );
}

// ============================================================
// PREVIEW FUNCTIONS
// ============================================================

function formatFirst5All() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  const tables = body.getTables();
  const listItems = body.getListItems();

  const results = { h1: 0, h2: 0, h3: 0, body: 0, lists: 0, tables: 0, images: 0 };

  log('\n=== Preview: Formatting first 5 of each type ===');

  for (const para of paragraphs) {
    if (results.h1 >= 5) break;
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
      let text = para.getText();
      if (!text.trim() || text.startsWith('<!--')) continue;

      text = text.replace(/^[▌▐│┃|]+\s*/g, '').trim();
      const newText = '▌ ' + text.toUpperCase();
      para.setText(newText);
      para.setFontFamily(STYLES.fonts.heading);
      para.setFontSize(STYLES.sizes.h1);
      para.setBold(true);

      const textEl = para.editAsText();
      textEl.setForegroundColor(0, 0, STYLES.colors.headingRed);
      if (newText.length > 2) textEl.setForegroundColor(2, newText.length - 1, STYLES.colors.charcoal);

      results.h1++;
    }
  }

  for (const para of paragraphs) {
    if (results.h2 >= 5) break;
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING2) {
      if (!para.getText().trim() || para.getText().startsWith('<!--')) continue;
      para.setFontFamily(STYLES.fonts.heading);
      para.setFontSize(STYLES.sizes.h2);
      para.setForegroundColor(STYLES.colors.darkGray3);
      para.setBold(true);
      results.h2++;
    }
  }

  for (const para of paragraphs) {
    if (results.h3 >= 5) break;
    if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING3) {
      if (!para.getText().trim() || para.getText().startsWith('<!--')) continue;
      para.setFontFamily(STYLES.fonts.heading);
      para.setFontSize(STYLES.sizes.h3);
      para.setForegroundColor(STYLES.colors.darkGray);
      para.setBold(true);
      results.h3++;
    }
  }

  for (const para of paragraphs) {
    if (results.body >= 5) break;
    if (para.getHeading() === DocumentApp.ParagraphHeading.NORMAL) {
      const text = para.getText();
      if (!text.trim() || text.startsWith('<!--') || text.startsWith('    ')) continue;
      const font = para.getFontFamily();
      if (font && font.toLowerCase().includes('mono')) continue;

      para.setFontFamily(STYLES.fonts.body);
      para.setFontSize(STYLES.sizes.body);
      para.setForegroundColor(STYLES.colors.charcoal);
      results.body++;
    }
  }

  for (let i = 0; i < Math.min(5, listItems.length); i++) {
    listItems[i].setAttributes({});
    listItems[i].setFontFamily(STYLES.fonts.body);
    listItems[i].setFontSize(STYLES.sizes.body);
    listItems[i].setForegroundColor(STYLES.colors.charcoal);
    listItems[i].setBold(false);
    listItems[i].setItalic(false);
    listItems[i].setSpacingBefore(0);
    listItems[i].setSpacingAfter(0);
    listItems[i].setLineSpacing(1.0);
    results.lists++;
  }

  for (let i = 0; i < Math.min(5, tables.length); i++) {
    formatSingleTable(tables[i]);
    results.tables++;
  }

  for (const para of paragraphs) {
    if (results.images >= 5) break;
    for (let i = 0; i < para.getNumChildren(); i++) {
      if (results.images >= 5) break;
      const child = para.getChild(i);
      if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
        const img = child.asInlineImage();
        const aspectRatio = img.getHeight() / img.getWidth();
        img.setWidth(STYLES.imageWidth);
        img.setHeight(Math.round(STYLES.imageWidth * aspectRatio));
        results.images++;
      }
    }
  }

  DocumentApp.getUi().alert('Preview Complete',
    `Formatted first 5 of each:\n\n` +
    `• H1: ${results.h1}\n• H2: ${results.h2}\n• H3: ${results.h3}\n` +
    `• Body: ${results.body}\n• Lists: ${results.lists}\n• Tables: ${results.tables}\n• Images: ${results.images}\n\n` +
    `Scroll through the document to review.`,
    DocumentApp.getUi().ButtonSet.OK);
}

function formatFirst5Tables() {
  const doc = DocumentApp.getActiveDocument();
  const tables = doc.getBody().getTables();
  const count = Math.min(5, tables.length);

  log('\n=== Preview: Formatting first 5 tables ===');

  for (let i = 0; i < count; i++) {
    formatSingleTable(tables[i]);
    log(`  Table ${i + 1} formatted`);
  }

  DocumentApp.getUi().alert('Preview Complete',
    `Formatted ${count} tables.\n\nScroll through to check styling.`,
    DocumentApp.getUi().ButtonSet.OK);
}

// ============================================================
// LOGGING UTILITY
// ============================================================

function log(message) {
  console.log(message);
  Logger.log(message);
}
