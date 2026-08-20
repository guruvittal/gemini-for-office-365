/**
 * Host Adapter Implementation for Microsoft Word (Word.run)
 * 
 * @author Sathya AG, Principal Architect, Google
 */
export class WordAdapter {
  constructor() {
    this.name = "Word";
  }

  // Read currently highlighted document text in Word
  async getSelectedText() {
    let selectedText = "";
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        if (selection.text && selection.text.trim()) {
          selectedText = selection.text.trim();
        }
      });
    } catch (err) {
      console.warn("Word selection read error:", err);
    }
    return selectedText;
  }

  // Read the entire document text for "Chat with this Document" Q&A
  async getFullDocumentText() {
    let fullText = "";
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        if (body.text && body.text.trim()) {
          fullText = body.text.trim();
        }
      });
    } catch (err) {
      console.warn("Word full document read error:", err);
    }
    return fullText;
  }

  // Water-tight insertion of rich HTML with smart In-Place Replacement via Word Content Controls
  async insertContent(htmlContent, mode = "smart") {
    if (!htmlContent) return;

    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = htmlContent;
    const finalHtml = tempDiv.innerHTML;

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");

      const contentControls = context.document.contentControls.getByTag("GEMINI_DRAFT");
      contentControls.load("items");
      await context.sync();

      const hasSelection = selection.text && selection.text.trim().length > 0;
      const hasDraft = contentControls.items && contentControls.items.length > 0;

      // 1. Priority 1: If user currently has text selected in Word, ALWAYS replace the current selection!
      if (hasSelection && mode !== "insert_cursor") {
        const insertedRange = selection.insertHtml(finalHtml, Word.InsertLocation.replace);
        try {
          const cc = insertedRange.insertContentControl();
          cc.tag = "GEMINI_DRAFT";
          cc.title = "Gemini AI Draft";
        } catch (tagErr) {
          console.warn("ContentControl tag note:", tagErr);
        }
        await context.sync();
        return;
      }

      // 2. Priority 2: If user clicked "Replace in Doc" without a new selection, update the previous draft in-place
      if ((mode === "replace_draft" || mode === "smart") && hasDraft) {
        try {
          const activeCc = contentControls.items[contentControls.items.length - 1];
          const targetRange = activeCc.getRange();
          activeCc.delete(false);
          const newRange = targetRange.insertHtml(finalHtml, Word.InsertLocation.replace);
          try {
            const newCc = newRange.insertContentControl();
            newCc.tag = "GEMINI_DRAFT";
            newCc.title = "Gemini AI Draft";
          } catch (e) {
            console.warn("ContentControl re-tag note:", e);
          }
          await context.sync();
          return;
        } catch (ccErr) {
          console.warn("ContentControl in-place replacement warning:", ccErr);
        }
      }

      // 3. Priority 3: Insert at cursor location and tag
      const insertedRange = selection.insertHtml(finalHtml, Word.InsertLocation.replace);
      try {
        const cc = insertedRange.insertContentControl();
        cc.tag = "GEMINI_DRAFT";
        cc.title = "Gemini AI Draft";
      } catch (tagErr) {
        console.warn("ContentControl tag note:", tagErr);
      }
      await context.sync();
    });
  }

  // Scan and execute in-document @gemini commands with Unique Content Marker Tracking
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    let targetPrompt = "";
    let markerId = "";

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const activeParagraphs = selection.paragraphs;
      activeParagraphs.load("items, text");

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items, text");
      await context.sync();

      if (!paragraphs.items || paragraphs.items.length === 0) {
        if (callbacks.onStatus) callbacks.onStatus("Status: No paragraphs in document");
        return;
      }

      const activeText = (activeParagraphs.items && activeParagraphs.items.length > 0 && activeParagraphs.items[0].text) 
        ? activeParagraphs.items[0].text.replace(/\u00A0/g, " ").replace(/[\r\n]/g, " ").trim() 
        : "";

      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const rawPText = p.text ? p.text.replace(/\u00A0/g, " ").replace(/[\r\n]/g, " ").trim() : "";

        const geminiIdx = rawPText.toLowerCase().indexOf("@gemini");
        if (geminiIdx !== -1) {
          const userPrompt = rawPText.substring(geminiIdx + 7).replace(/^[:\s]+/, "").trim();

          const isCursorInThisParagraph = (activeText === rawPText);
          const isReadyToExecute = (!isCursorInThisParagraph || forceRun) && userPrompt.length >= 2;

          if (isReadyToExecute) {
            targetPrompt = userPrompt;
            markerId = "GEMINI_MARKER_" + Date.now();

            if (callbacks.onStatus) callbacks.onStatus(`Found: "@gemini ${userPrompt.substring(0, 20)}..."`);

            p.insertText(`⚡ [Gemini Enterprise is generating response... ${markerId}]`, Word.InsertLocation.replace);
            await context.sync();
            break;
          } else if (isCursorInThisParagraph && !forceRun) {
            if (callbacks.onStatus) callbacks.onStatus(`Typing: "@gemini ${userPrompt.substring(0, 15)}..." (press ENTER)`);
          }
        }
      }
    });

    if (targetPrompt && markerId && callbacks.executePrompt) {
      const formattedHtml = await callbacks.executePrompt(targetPrompt);
      if (formattedHtml) {
        await Word.run(async (context) => {
          const searchResults = context.document.body.search(markerId, { matchCase: true });
          searchResults.load("items");
          await context.sync();

          if (searchResults.items && searchResults.items.length > 0) {
            const markerItem = searchResults.items[0];
            const targetParagraphs = markerItem.paragraphs;
            targetParagraphs.load("items");
            await context.sync();

            if (targetParagraphs.items && targetParagraphs.items.length > 0) {
              targetParagraphs.items[0].insertHtml(formattedHtml, Word.InsertLocation.replace);
              await context.sync();
            }
          }
        });
      }
    }
  }
}
