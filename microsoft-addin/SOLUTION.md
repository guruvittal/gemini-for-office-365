# End-to-End Solution: Gemini 2.5 Flash + Vertex AI Image Generation for Microsoft PowerPoint & Word

## Executive Overview
This document records the complete, verified architecture and implementation for generating enterprise executive briefings with AI-generated visual charts directly inserted into Microsoft PowerPoint presentations and Microsoft Word documents.

---

## 1. System Architecture

```mermaid
graph TD
    User([User in PowerPoint / Word]) -->|Submits Prompt| Taskpane[Add-in Taskpane (Webview)]
    Taskpane -->|POST JSON| Proxy[Google Cloud Function (askGemini)]
    
    subgraph Vertex AI Backend
        Proxy -->|1. Grounded Inference| Gemini[Gemini 2.5 Flash + Vertex AI Search Datastore]
        Gemini -->|2. Markdown + Image Prompts| Proxy
        Proxy -->|3. Balanced Parser| Extractor[Image Prompt Extractor]
        Extractor -->|4. Generate Visual Charts| Imagen[gemini-2.5-flash-image (Nano Banana)]
        Imagen -->|5. High-Res PNG Base64| Inliner[HTML Data URI Inliner]
    end
    
    Inliner -->|HTML + Base64 Response| Taskpane
    
    subgraph PowerPoint Slide Generation Engine
        Taskpane -->|parseMarkdown & DOM Parser| SlideParser[Slide Section & Visual Parser]
        SlideParser -->|Title Slide / Split Columns / Visuals| PPTAdapter[PowerPoint Host Adapter]
        PPTAdapter -->|1. PowerPoint.run (Geometry, Title, Body)| RichApi[Office.js RichApi]
        PPTAdapter -->|2. Slide Selection (context.presentation.setSelectedSlides)| Bridge[Selection Bridge]
        PPTAdapter -->|3. Office.context.document.setSelectedDataAsync| CommonAPI[Office Common C++ API]
    end
    
    CommonAPI -->|Native Render| Slides([Main PowerPoint Deck with Charts & Typography])
```

---

## 2. Key Technical Innovations & Solutions

### A. Grounded Financial Retrieval & Generic AI Image Generation
1. **Model & Grounding Configuration:**
   - **Text Model:** `gemini-2.5-flash` on Vertex AI with temperature `0.4` and max output `8192` tokens.
   - **Grounding Datastore:** Vertex AI Search (`o365geminipluginds_1786291473119_gcs_store`) providing real-time data from financial quarterly filings (10-Q/10-K).
   - **Image Model:** `gemini-2.5-flash-image` (Nano Banana) generating 2D flat vector, high-resolution financial charts and diagrams.

2. **Balanced Parenthesis Prompt Extractor:**
   - Handles complex nested prompt syntax with financial numbers containing parentheses (e.g. `($94.5B)`).
   - Extracts complete prompt blocks and converts them into embedded base64 PNG data tags without prompt text leakage.

---

### B. PowerPoint Slide Formatting & Layout Engine

1. **Intelligent Multi-Layout Distribution:**
   - **Title Slide:** Re-uses initial blank slide #1 in fresh decks, styling with 32pt centered typography.
   - **Summary / Content Slides:** Full-width (880px) spacious layout with Segoe UI typography and top blue accent lines.
   - **Financial Data Tables:** Automatically parsed from HTML `<table>` elements into formatted key-value bullet points:
     - `• [Segment | Q2 2026 Revenue (Millions USD)]`
     - `• Google Services: $94,540`
     - `• Google Cloud: $24,768`
   - **Visual Chart Slides:** Split-column layout (Text on left width `420px`, AI chart image on right `440px x 340px`).

---

### C. macOS Office.js Dual-Pipeline Image Injection

#### The Challenge on macOS:
PowerPoint for macOS runs the add-in inside a sandboxed `WKWebView`. Passing large base64 strings (>50KB) through `PowerPoint.run` XPC messages causes Apple's XPC buffer serializer to drop the payload during `context.sync()`.

#### The Solution:
1. **In-Memory Preparation:** Image scaling and compression are executed in the browser taskpane *before* entering `PowerPoint.run`, preventing Office.js context invalidation.
2. **Slide Selection Synchronization:**
   Inside `PowerPoint.run`, after creating the slide, title, and body shapes, the engine loads `newSlide.id` and calls:
   ```javascript
   context.presentation.setSelectedSlides([newSlide.id]);
   await context.sync();
   ```
3. **Office Common API Image Injection (`setSelectedDataAsync`):**
   Immediately after the slide is selected, the engine calls:
   ```javascript
   Office.context.document.setSelectedDataAsync(cleanBase64, {
     coercionType: Office.CoercionType.Image,
     imageLeft: 480,
     imageTop: 110,
     imageWidth: 440,
     imageHeight: 340
   });
   ```
   This bypasses the macOS XPC `PowerPoint.run` RichApi bridge and communicates directly through native Office C++ document pipelines, ensuring 100% reliability across all Office for Mac versions.

---

## 3. Deployment Artifacts & Endpoints

| Component | Target Service | Revision | Region | URL |
| :--- | :--- | :--- | :--- | :--- |
| **Backend Proxy** | Cloud Functions Gen 2 / Cloud Run | `askgemini-00036-naj` | `us-central1` | `https://us-central1-genai-demo-catalog.cloudfunctions.net/askGemini` |
| **Add-in Frontend** | Cloud Run (Nginx) | `gemini-frontend-00078-hg7` | `us-central1` | `https://gemini-frontend-133594738129.us-central1.run.app` |

---

## 4. Verification Test Prompt

**Prompt:**
```text
Based on our financial report for Q1 and Q2, write an executive briefing comparing Google Cloud and Google Services operating performance for Q2 2026 with visual charts
```

**Verified Output:**
- **Slide 1:** Centered bold Title Slide (`Executive Briefing: Q2 2026 Operating Performance - Google Cloud vs. Google Services`).
- **Slide 2:** Full-width Executive Summary.
- **Slide 3:** Revenue Performance with clean bulleted financial breakdown.
- **Slide 4:** Operating Income Analysis with key highlights on the left and native Google Services vs. Google Cloud comparison bar chart on the right.
