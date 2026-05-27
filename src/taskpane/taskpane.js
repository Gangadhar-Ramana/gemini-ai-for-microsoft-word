/*
 * Gemini AI for Office - Task Pane Implementation
 * Author: Anson Lai
 * Location: Vancouver, Canada
 * Description: Word add-in integrating Google Gemini AI for document editing and analysis
 */

/* global document, Office, Word, localStorage */

import { marked } from 'marked';
import { diff_match_patch } from 'diff-match-patch';
import "./taskpane.css";

import {
  registerChatUiHandlers,
  setupScrollToBottom,
  createTypingIndicator,
  shakeInput,
  addMessageToChat,
  updateSystemMessage,
  addRetryButton,
  hideAllRetryButtons,
  removeMessage
} from './modules/chat/chat-ui.js';
import {
  maintainHistoryWindow,
  validateHistoryPairs,
  sanitizeHistory,
  removeAllFunctionPairs,
  createFreshStartWithContext
} from './modules/chat/chat-history.js';
import {
  initAgenticTools,
  executeRedline,
  executeComment,
  executeHighlight,
  executeNavigate,
  executeResearch,
  executeInsertListItem,
  executeEditList,
  executeConvertHeadersToList,
  executeEditTable,
  executeEditSection
} from './modules/commands/agentic-tools.js';
import { setPlatform, wrapInDocumentFragment } from '@ansonlai/docx-redline-js';

// Configure marked for GFM (GitHub Flavored Markdown) with tables, breaks, etc.
marked.setOptions({
  gfm: true,           // Enable GitHub Flavored Markdown
  breaks: true,        // Convert \n to <br>
});

// ==================== CONFIGURATION CONSTANTS ====================

const DEFAULT_AUTHOR = "Gemini AI";
const GLANCE_COLLAPSED_STORAGE_KEY = "glanceCollapsed";
const GOOGLE_MODEL_LIST_STORAGE_KEY = "geminiUsableGoogleModels";
const MODEL_DEFAULT_MIGRATION_KEY = "gemini25DefaultModelMigrationApplied";
const DEFAULT_FAST_MODEL = "gemini-2.5-flash";
const DEFAULT_SLOW_MODEL = "gemini-2.5-pro";
const EDIT_FALLBACK_MODEL = "gemini-2.5-flash";

const DEFAULT_MODEL_OPTIONS = [
  { id: "deep-research-max-preview-04-2026", label: "Deep Research Max Preview (Apr-21-2026)", method: "generateContent" },
  { id: "deep-research-preview-04-2026", label: "Deep Research Preview (Apr-21-2026)", method: "generateContent" },
  { id: "deep-research-pro-preview-12-2025", label: "Deep Research Pro Preview (Dec-12-2025)", method: "generateContent" },
  { id: "gemini-2.0-flash", label: "Gemini 2.0 Flash", method: "generateContent" },
  { id: "gemini-2.0-flash-001", label: "Gemini 2.0 Flash 001", method: "generateContent" },
  { id: "gemini-2.0-flash-lite", label: "Gemini 2.0 Flash-Lite", method: "generateContent" },
  { id: "gemini-2.0-flash-lite-001", label: "Gemini 2.0 Flash-Lite 001", method: "generateContent" },
  { id: "gemini-2.5-computer-use-preview-10-2025", label: "Gemini 2.5 Computer Use Preview 10-2025", method: "generateContent" },
  { id: "gemini-2.5-flash", label: "Gemini 2.5 Flash", method: "generateContent" },
  { id: "gemini-2.5-flash-image", label: "Nano Banana", method: "generateContent" },
  { id: "gemini-2.5-flash-lite", label: "Gemini 2.5 Flash-Lite", method: "generateContent" },
  { id: "gemini-2.5-flash-preview-tts", label: "Gemini 2.5 Flash Preview TTS", method: "generateContent" },
  { id: "gemini-2.5-pro", label: "Gemini 2.5 Pro", method: "generateContent" },
  { id: "gemini-2.5-pro-preview-tts", label: "Gemini 2.5 Pro Preview TTS", method: "generateContent" },
  { id: "gemini-3.1-flash-image-preview", label: "Nano Banana 2", method: "generateContent" },
  { id: "gemini-3.1-flash-lite", label: "Gemini 3.1 Flash-Lite", method: "generateContent" },
  { id: "gemini-3.1-flash-lite-preview", label: "Gemini 3.1 Flash Lite Preview", method: "generateContent" },
  { id: "gemini-3.1-flash-tts-preview", label: "Gemini 3.1 Flash TTS Preview", method: "generateContent" },
  { id: "gemini-3.1-pro-preview", label: "Gemini 3.1 Pro Preview", method: "generateContent" },
  { id: "gemini-3.1-pro-preview-customtools", label: "Gemini 3.1 Pro Preview Custom Tools", method: "generateContent" },
  { id: "gemini-3-flash-preview", label: "Gemini 3 Flash Preview", method: "generateContent" },
  { id: "gemini-3-pro-image-preview", label: "Nano Banana Pro", method: "generateContent" },
  { id: "gemini-3-pro-preview", label: "Gemini 3 Pro Preview", method: "generateContent" },
  { id: "gemini-flash-latest", label: "Gemini Flash Latest", method: "generateContent" },
  { id: "gemini-flash-lite-latest", label: "Gemini Flash-Lite Latest", method: "generateContent" },
  { id: "gemini-pro-latest", label: "Gemini Pro Latest", method: "generateContent" },
  { id: "gemini-robotics-er-1.5-preview", label: "Gemini Robotics-ER 1.5 Preview", method: "generateContent" },
  { id: "gemini-robotics-er-1.6-preview", label: "Gemini Robotics-ER 1.6 Preview", method: "generateContent" },
  { id: "gemma-4-26b-a4b-it", label: "Gemma 4 26B A4B IT", method: "generateContent" },
  { id: "gemma-4-31b-it", label: "Gemma 4 31B IT", method: "generateContent" },
  { id: "lyria-3-clip-preview", label: "Lyria 3 Clip Preview", method: "generateContent" },
  { id: "lyria-3-pro-preview", label: "Lyria 3 Pro Preview", method: "generateContent" },
  { id: "nano-banana-pro-preview", label: "Nano Banana Pro", method: "generateContent" },
  { id: "gemini-2.5-flash-native-audio-latest", label: "Gemini 2.5 Flash Native Audio Latest", method: "bidiGenerateContent" },
  { id: "gemini-2.5-flash-native-audio-preview-09-2025", label: "Gemini 2.5 Flash Native Audio Preview 09-2025", method: "bidiGenerateContent" },
  { id: "gemini-2.5-flash-native-audio-preview-12-2025", label: "Gemini 2.5 Flash Native Audio Preview 12-2025", method: "bidiGenerateContent" },
  { id: "gemini-3.1-flash-live-preview", label: "Gemini 3.1 Flash Live Preview", method: "bidiGenerateContent" }
];

// Safety settings for Gemini API (disable all safety blocks)
const SAFETY_SETTINGS_BLOCK_NONE = [
  { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
  { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
  { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
  { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
];

// Search and text limits
const SEARCH_LIMITS = {
  MAX_LENGTH: 100,           // Max search string length for comments/highlights
  MAX_LENGTH_MODIFY: 80,     // Max search string length for modify_text operations
  SUFFIX_LENGTH: 60,         // Suffix length for range expansion
  RETRY_LENGTH: 30           // Fallback shorter search length for retries
};

// Document processing limits
const DOCUMENT_LIMITS = {
  MAX_WORDS: 30000,          // Approx 40 pages, ~40k tokens
  MAX_LOOPS: 6,              // Maximum tool execution loops
  MAX_NO_PROGRESS_TOOL_LOOPS: 2, // Stop when the same mutation tool cycle keeps applying 0 changes
  TOKEN_MULTIPLIER: 1.33     // Words to tokens conversion factor
};

// Storage quotas
const STORAGE_LIMITS = {
  SAFE_LIMIT: 4500000,       // ~4.5MB safe limit for localStorage
  MIN_PRUNE_COUNT: 5         // Minimum checkpoints to prune when quota exceeded
};

// API generation limits
const API_LIMITS = {
  MAX_OUTPUT_TOKENS: 48000   // Maximum tokens for AI response output
};

// Timeout limits for API calls
const TIMEOUT_LIMITS = {
  FETCH_TIMEOUT_MS: 60000,        // 60s timeout per individual API call
  TOTAL_REQUEST_TIMEOUT_MS: 180000 // 3 min total timeout for entire request (including tool loops)
};

// Global abort controller for cancelling requests
let currentRequestController = null;
/**
 * Extracts enhanced document context with rich formatting metadata.
 * Returns an object with enhanced paragraph notation and section mapping.
 * 
 * Format: [P#|Style|ListInfo|TableInfo|SectionInfo] Text
 * Examples:
 *   [P1|Normal] Regular paragraph
 *   [P2|Heading1] Chapter heading
 *   [P3|ListNumber|L1:0|§] 1. Section header (starts section 1)
 *   [P4|Normal|§1] Body text belonging to section 1
 *   [P5|Normal|T:1,2] Table cell at row 1, column 2
 */
async function extractEnhancedDocumentContext(context) {
  const body = context.document.body;
  const paragraphs = body.paragraphs;

  // Load all relevant paragraph properties
  paragraphs.load("items");
  await context.sync();

  // Load detailed properties for each paragraph
  for (const para of paragraphs.items) {
    para.load("text, style, listItemOrNullObject, parentTableOrNullObject, parentTableCellOrNullObject");
  }
  await context.sync();

  // Load list details for paragraphs that are list items
  for (const para of paragraphs.items) {
    if (!para.listItemOrNullObject.isNullObject) {
      para.listItemOrNullObject.load("level, listString");
    }
    if (!para.parentTableCellOrNullObject.isNullObject) {
      para.parentTableCellOrNullObject.load("rowIndex, cellIndex");
    }
  }

  await context.sync();

  // Build enhanced paragraph data
  const enhancedParagraphs = [];
  let currentSection = null;      // Current section number (e.g., "1", "2")
  let currentSubSection = null;   // Current subsection (e.g., "1.1", "2.3")
  let sectionCounter = 0;         // Tracks top-level sections
  let lastListLevel = -1;         // Tracks list nesting level
  let sectionStack = [];          // Stack for tracking nested sections

  for (let i = 0; i < paragraphs.items.length; i++) {
    const para = paragraphs.items[i];
    const text = para.text || "";
    const style = para.style || "Normal";

    // Build metadata parts
    const metaParts = [style];

    // Check if paragraph is a list item
    let isListItem = false;
    let listLevel = -1;
    let listString = "";

    if (!para.listItemOrNullObject.isNullObject) {
      isListItem = true;
      listLevel = para.listItemOrNullObject.level || 0;
      listString = para.listItemOrNullObject.listString || "";

      // Determine list type from style name
      const isNumbered = style.toLowerCase().includes("number") ||
        style.toLowerCase().includes("list number") ||
        /^\d+[.)]/.test(listString);
      const listType = isNumbered ? "ListNumber" : "ListBullet";

      // Replace style with more specific list type
      metaParts[0] = listType;

      // Add list ID and level (using a simple counter-based ID)
      metaParts.push(`L:${listLevel}`);
    }

    // Check if paragraph is in a table
    let isInTable = false;
    if (!para.parentTableCellOrNullObject.isNullObject) {
      isInTable = true;
      const rowIndex = para.parentTableCellOrNullObject.rowIndex || 0;
      const cellIndex = para.parentTableCellOrNullObject.cellIndex || 0;
      metaParts.push(`T:${rowIndex},${cellIndex}`);
    }

    // Section detection for legal contract patterns
    let sectionMarker = "";

    if (isListItem && !isInTable) {
      // This list item could be a section header
      // Detect section headers: list items at level 0 or items that start new sections

      if (listLevel === 0) {
        // Top-level list item = new section
        sectionCounter++;
        currentSection = String(sectionCounter);
        currentSubSection = null;
        sectionStack = [currentSection];
        sectionMarker = "§";  // Mark as section header
        lastListLevel = listLevel;
      } else if (listLevel > lastListLevel) {
        // Nested list item = subsection
        const parentSection = sectionStack[sectionStack.length - 1] || currentSection;
        const subNum = sectionStack.length;
        currentSubSection = `${parentSection}.${listLevel}`;
        sectionStack.push(currentSubSection);
        sectionMarker = "§";  // Also mark as subsection header
        lastListLevel = listLevel;
      } else if (listLevel <= lastListLevel && listLevel > 0) {
        // Same or shallower nested level - pop stack and create new subsection
        while (sectionStack.length > listLevel + 1) {
          sectionStack.pop();
        }
        const parentSection = sectionStack[0] || currentSection;
        currentSubSection = `${parentSection}.${listLevel}`;
        sectionStack[listLevel] = currentSubSection;
        sectionMarker = "§";
        lastListLevel = listLevel;
      }

      if (sectionMarker) {
        metaParts.push(sectionMarker);
      }
    } else if (!isListItem && !isInTable && currentSection) {
      // Non-list paragraph following a section header = section body
      const belongsTo = currentSubSection || currentSection;
      metaParts.push(`§${belongsTo}`);
    }

    // Build the enhanced notation
    const metaString = metaParts.join("|");
    const enhancedLine = `[P${i + 1}|${metaString}] ${text}`;

    enhancedParagraphs.push({
      index: i + 1,
      text: text,
      style: style,
      isListItem: isListItem,
      listLevel: listLevel,
      isInTable: isInTable,
      section: currentSection,
      subSection: currentSubSection,
      isSectionHeader: sectionMarker === "§",
      enhancedLine: enhancedLine
    });
  }

  return {
    paragraphs: enhancedParagraphs,
    formattedText: enhancedParagraphs.map(p => p.enhancedLine).join("\n"),
    sectionCount: sectionCounter
  };
}

let chatHistory = [];
let toolsExecutedInCurrentRequest = [];  // Track successful tool executions for recovery

function generateSuccessMessage(executedTools = []) {
  const successfulTools = (Array.isArray(executedTools) ? executedTools : [])
    .filter(tool => tool && tool.success !== false);

  if (successfulTools.length === 0) {
    return "";
  }

  const latestTool = successfulTools[successfulTools.length - 1];
  const resultText = String(latestTool.result || "").trim();
  if (resultText) {
    return resultText;
  }

  const toolName = latestTool.name || "";
  if (toolName === "insert_word_equation") return "Inserted the Word equation.";
  if (toolName === "convert_text_to_word_math") return "Converted the requested text to inline Word math.";
  if (toolName === "format_text_occurrences") return "Applied the requested text formatting.";
  if (toolName === "highlight_text") return "Highlighted the requested text.";
  if (toolName === "insert_comment") return "Inserted the requested comment.";
  if (toolName === "apply_redlines") return "Applied the requested document edits.";
  if (toolName === "run_word_script") return "Applied the requested Word action plan.";
  if (toolName === "edit_list" || toolName === "insert_list_item" || toolName === "convert_headers_to_list") {
    return "Updated the requested list formatting.";
  }
  if (toolName === "edit_table") return "Updated the requested table.";
  if (toolName === "edit_section") return "Updated the requested section.";

  return "Task completed successfully.";
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    setPlatform(Office?.context?.platform);
    migrateDefaultModelSelections();
    document.getElementById("sideload-msg").style.display = "none";
    // Show main view by default
    showMainView();

    // Add event listener for the chat send button (Fast)
    document.getElementById("send-button").onclick = () => sendChatMessage('fast');

    // Add event listener for the THINK button (Slow)
    document.getElementById("think-button").onclick = () => sendChatMessage('slow');

    // Add Enter key support for chat (Shift+Enter for new line)
    document.getElementById("chat-input").addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        if (e.shiftKey) {
          // Shift+Enter: New line (default behavior)
          return;
        }
        e.preventDefault();
        if (e.ctrlKey || e.metaKey) {
          // Ctrl+Enter or Cmd+Enter: Thinking chat (slow)
          sendChatMessage('slow');
        } else {
          // Enter: Regular chat (fast)
          sendChatMessage('fast');
        }
      }
    });

    // Add event listeners for settings UI
    document.getElementById("settings-button").onclick = showSettingsView;
    document.getElementById("save-api-key").onclick = saveApiKey;
    document.getElementById("back-to-main").onclick = showMainView;
    document.getElementById("refresh-google-models-button").onclick = refreshGoogleModels;

    // Add event listener for refresh chat button
    document.getElementById("refresh-chat-button").onclick = refreshChat;

    // Add event listener for Glance refresh
    document.getElementById("refresh-glance-button").onclick = runGlanceChecks;
    document.getElementById("toggle-glance-button").onclick = () => {
      const container = document.getElementById("glance-container");
      if (!container) return;
      const shouldCollapse = !container.classList.contains("collapsed");
      saveGlanceCollapsedState(shouldCollapse);
      applyGlanceCollapsedState(shouldCollapse);
    };

    // Add event listener for Add Glance Card
    document.getElementById("add-glance-card-button").onclick = () => {
      const settings = loadGlanceSettings();
      settings.push({
        id: 'q' + Date.now(),
        title: 'New Question',
        question: 'What would you like to check?'
      });
      saveGlanceSettings(settings);
      renderGlanceSettings();
    };

    // Check for API key on load
    if (!loadApiKey()) {
      showWelcomeScreen();
    } else {
      // Run Glance checks if key exists
      renderGlanceMain();
      runGlanceChecks();
    }

    // Accordion Event Listeners
    setupAccordion("glance-settings-header", "glance-settings-content");
    setupAccordion("advanced-settings-header", "advanced-settings-content");

    // Scroll-to-bottom button setup
    setupScrollToBottom();

    // Add event listener for refresh author button
    document.getElementById("refresh-author-button").onclick = async () => {
      const author = await fetchDocumentAuthor();
      if (author) {
        document.getElementById("redline-author-input").value = author;
        saveRedlineAuthor(author);
      }
    };

    // Add event listeners for Redline settings
    document.getElementById("redline-toggle").onchange = (e) => {
      saveRedlineSetting(e.target.checked);
    };

    document.getElementById("redline-author-input").oninput = (e) => {
      saveRedlineAuthor(e.target.value);
    };

    // Update checkpoint status on load (internal only now)
    // updateCheckpointStatus(); // UI removed, but we can keep tracking internally if needed, or just remove this call.
  }
});

function showWelcomeScreen() {
  const chatMessages = document.getElementById("chat-messages");
  chatMessages.innerHTML = ""; // Clear existing messages

  const welcomeContainer = document.createElement("div");
  welcomeContainer.className = "welcome-container";

  welcomeContainer.innerHTML = `
    <div class="welcome-header">
      <h2>Get Started in 30 Seconds</h2>
    </div>
    <div class="welcome-step">
      <div class="step-number">1</div>
      <div class="step-content">
        <p>Go to <a href="https://aistudio.google.com/app/api-keys" target="_blank">Google AI Studio</a>.</p>
      </div>
    </div>
    <div class="welcome-step">
      <div class="step-number">2</div>
      <div class="step-content">
        <p>Click <strong>Create API key</strong> (top left).</p>
      </div>
    </div>
    <div class="welcome-step">
      <div class="step-number">3</div>
      <div class="step-content">
        <p>Select your project (or create new) and copy the Google API key string.</p>
      </div>
    </div>
    <div class="welcome-step">
      <div class="step-number">4</div>
      <div class="step-content">
        <p>Click the <strong>Gear Icon</strong> <span style="font-size: 1.2em;">&#9881;</span> at the top right corner to enter your key.</p>
      </div>
    </div>
    <div class="welcome-note">
      <p style="text-align: right;">The free tier is <em>plenty</em> for personal use.</p>
    </div>

    <hr class="welcome-divider">

    <div class="welcome-header">
      <h2 >Features</h2>
    </div>

    <div class="feature-explanation">
      <h3>Document Tools</h3>
      <p>Chat with an assistant who can access to tools that can <strong>edit text</strong>, <strong>search Google</strong>, <strong>highlight key info</strong>, and <strong>leave comments</strong>.  These tools allow the assistant to interact with your document naturally and help you with your tasks.</p>
    </div>

    <div class="feature-explanation">
      <h3>Glance Checks</h3>
      <p>Set up custom criteria (like <em>Grammar</em> or <em>Factual Accuracy</em>) to automatically check every document you open.  You can customize these questions in Settings.</p>
    </div>

    <div class="feature-explanation">
      <h3>System Prompts</h3>
      <p>Customize how the AI behaves. You can tell it to be a <em>Grade 10 student working on an English paper</em> or an <em>associate lawyer at a New York law firm specializing in contracts</em>.  Give it context and instructions you think would be helpful.</p>
    </div>

    <div class="feature-explanation">
      <h3>Model Choices</h3>
      <p><strong>Fast Model:</strong> This model is used for regular chats and is great for quick edits and simple questions.  It is fast and cheap.</p>
      <p><strong>Slow Model:</strong> This model is used when you select "Think".  It provides deep analysis and basic online searches.  It is slower and more expensive, but provides more thorough results.</p>
    </div>

    <div class="welcome-footer">
      <p><em>If you have any questions, please reach out to us at <a href="mailto:support@reference.legal">support@reference.legal</a>.</em></p>
    </div>
  `;

  chatMessages.appendChild(welcomeContainer);
}

// --- Settings & View Management ---

function switchView(hideId, showId) {
  const hideEl = document.getElementById(hideId);
  const showEl = document.getElementById(showId);

  if (!hideEl || !showEl) return;

  // Fade out current
  hideEl.classList.add("view-hidden");
  hideEl.classList.remove("view-container"); // Ensure it doesn't conflict

  setTimeout(() => {
    hideEl.style.display = "none";
    showEl.style.display = "block";

    // Force reflow
    void showEl.offsetWidth;

    // Fade in new
    showEl.classList.remove("view-hidden");
    showEl.classList.add("view-container");
  }, 200); // Match CSS transition speed
}

function showSettingsView() {
  document.getElementById("settings-button").style.display = "none";
  document.getElementById("refresh-chat-button").style.display = "none";

  switchView("main-view", "settings-view");

  // Load current key into input
  const currentKey = loadApiKey();
  if (currentKey) {
    document.getElementById("api-key-input").value = currentKey;
  }
  // Load current models
  const currentFastModel = loadModel('fast');
  populateModelSelect("model-select-fast", currentFastModel);
  if (currentFastModel) {
    document.getElementById("model-select-fast").value = currentFastModel;
  }
  const currentSlowModel = loadModel('slow');
  populateModelSelect("model-select-slow", currentSlowModel);
  if (currentSlowModel) {
    document.getElementById("model-select-slow").value = currentSlowModel;
  }
  // Load current system message
  const currentSystemMessage = loadSystemMessage();
  if (currentSystemMessage) {
    document.getElementById("system-message-input").value = currentSystemMessage;
  }
  // Render Glance settings
  renderGlanceSettings();

  // Load redline setting
  const redlineEnabled = loadRedlineSetting();
  document.getElementById("redline-toggle").checked = redlineEnabled;

  // Load redline author setting
  const redlineAuthor = loadRedlineAuthor();
  document.getElementById("redline-author-input").value = redlineAuthor;
}

function showMainView() {
  document.getElementById("settings-button").style.display = "block";
  document.getElementById("refresh-chat-button").style.display = "block";

  switchView("settings-view", "main-view");

  renderGlanceMain();
}


function refreshChat() {
  // Cancel any ongoing request
  if (currentRequestController) {
    currentRequestController.abort();
    currentRequestController = null;
    console.log("Active request cancelled by refresh.");
  }

  // Immediately unlock UI (in case it was locked by an active request)
  const chatInput = document.getElementById("chat-input");
  const sendButton = document.getElementById("send-button");
  const thinkButton = document.getElementById("think-button");

  if (chatInput) {
    chatInput.disabled = false;
    chatInput.value = "";
    chatInput.focus();
  }
  if (sendButton) sendButton.disabled = false;
  if (thinkButton) thinkButton.disabled = false;

  // Clear chat history
  chatHistory = [];

  // Clear the chat messages UI
  const chatMessages = document.getElementById("chat-messages");
  chatMessages.innerHTML = "";

  // Add the welcome message back
  const welcomeMessage = document.createElement("div");
  welcomeMessage.className = "chat-message system";
  welcomeMessage.textContent = "Welcome! Ask me to assist you in editing this document.";
  chatMessages.appendChild(welcomeMessage);

  // Add a system message confirming the refresh
  addMessageToChat("System", "Chat history cleared. Starting new conversation.");
}

function saveApiKey() {
  const apiKey = document.getElementById("api-key-input").value;
  const fastModel = document.getElementById("model-select-fast").value;
  const slowModel = document.getElementById("model-select-slow").value;
  const systemMessage = document.getElementById("system-message-input").value;
  const redlineEnabled = document.getElementById("redline-toggle").checked;
  const redlineAuthor = document.getElementById("redline-author-input").value;

  if (apiKey && apiKey.trim() !== "") {
    localStorage.setItem("geminiApiKey", apiKey);
    localStorage.setItem("geminiModelFast", fastModel);
    localStorage.setItem("geminiModelSlow", slowModel);
    localStorage.setItem("geminiSystemMessage", systemMessage);
    saveRedlineSetting(redlineEnabled);
    saveRedlineAuthor(redlineAuthor);
    // Glance settings are saved automatically on change
    showMainView();
    addMessageToChat("System", "Settings saved successfully.");
    // Re-run checks with new settings
    runGlanceChecks();
  } else {
    addMessageToChat("System", "API Key cannot be empty.");
  }
}

function loadApiKey() {
  // First check localStorage (user-provided key takes precedence)
  const storedKey = localStorage.getItem("geminiApiKey");
  if (storedKey && storedKey.trim() !== "") {
    return storedKey;
  }
}

function migrateDefaultModelSelections() {
  if (localStorage.getItem(MODEL_DEFAULT_MIGRATION_KEY) === "true") return;
  const fastModel = normalizeModelName(localStorage.getItem("geminiModelFast"));
  if (fastModel === "gemini-3.1-flash-lite") {
    localStorage.setItem("geminiModelFast", DEFAULT_FAST_MODEL);
  }
  localStorage.setItem(MODEL_DEFAULT_MIGRATION_KEY, "true");
}

function normalizeModelName(modelName) {
  return String(modelName || "").trim().replace(/^models\//, "");
}

function isLiveModel(modelName) {
  return /live/i.test(normalizeModelName(modelName));
}

function loadGoogleModelOptions() {
  const merged = new Map(DEFAULT_MODEL_OPTIONS.map(model => [model.id, model]));
  const stored = localStorage.getItem(GOOGLE_MODEL_LIST_STORAGE_KEY);
  if (stored) {
    try {
      const parsed = JSON.parse(stored);
      if (Array.isArray(parsed)) {
        for (const model of parsed) {
          if (model?.id) merged.set(model.id, model);
        }
      }
    } catch (error) {
      console.warn("Unable to parse stored Google model list:", error);
    }
  }
  return [...merged.values()].sort((a, b) => a.label.localeCompare(b.label));
}

function populateModelSelect(selectId, selectedModel) {
  const select = document.getElementById(selectId);
  if (!select) return;
  const selected = normalizeModelName(selectedModel);
  select.innerHTML = "";

  for (const model of loadGoogleModelOptions()) {
    const option = document.createElement("option");
    option.value = model.id;
    option.textContent = `${model.label}${model.method === "bidiGenerateContent" ? " (Live API)" : ""}`;
    if (model.id === selected) option.selected = true;
    select.appendChild(option);
  }
}

async function refreshGoogleModels() {
  const status = document.getElementById("google-model-refresh-status");
  const button = document.getElementById("refresh-google-models-button");
  const apiKey = document.getElementById("api-key-input")?.value?.trim() || loadApiKey();
  if (!apiKey) {
    if (status) status.textContent = "Enter and save your Gemini API key first.";
    return;
  }

  if (button) button.disabled = true;
  if (status) status.textContent = "Checking models available to this API key...";

  try {
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data.error?.message || `ListModels failed with HTTP ${response.status}`);
    }

    const usable = (data.models || [])
      .map(model => ({
        id: normalizeModelName(model.name),
        label: model.displayName || normalizeModelName(model.name),
        methods: model.supportedGenerationMethods || []
      }))
      .filter(model => model.methods.includes("generateContent") || model.methods.includes("bidiGenerateContent"))
      .map(model => ({
        id: model.id,
        label: model.label,
        method: model.methods.includes("generateContent") ? "generateContent" : "bidiGenerateContent",
        methods: model.methods
      }));

    if (usable.length === 0) {
      throw new Error("No generateContent or Live API models were returned for this key.");
    }

    localStorage.setItem(GOOGLE_MODEL_LIST_STORAGE_KEY, JSON.stringify(usable));
    populateModelSelect("model-select-fast", loadModel("fast"));
    populateModelSelect("model-select-slow", loadModel("slow"));

    const textModels = usable.filter(model => model.methods.includes("generateContent")).length;
    const liveModels = usable.filter(model => model.methods.includes("bidiGenerateContent")).length;
    if (status) {
      status.textContent = `Found ${usable.length} usable model(s): ${textModels} normal, ${liveModels} Live API.`;
    }
  } catch (error) {
    console.error("Unable to list Google models:", error);
    if (status) status.textContent = `Unable to list models: ${error.message || String(error)}`;
  } finally {
    if (button) button.disabled = false;
  }
}

function loadModel(type = 'fast') {
  const key = type === 'slow' ? "geminiModelSlow" : "geminiModelFast";
  const storedModel = localStorage.getItem(key);
  if (storedModel && storedModel.trim() !== "") {
    return normalizeModelName(storedModel);
  }
  // Defaults
  return type === 'slow' ? DEFAULT_SLOW_MODEL : DEFAULT_FAST_MODEL;
}

function loadSystemMessage() {
  const storedMessage = localStorage.getItem("geminiSystemMessage");
  if (storedMessage && storedMessage.trim() !== "") {
    return storedMessage;
  }
  return "Example: You are assisting an undergraduate student with their academic paper. You must be specific, precise, and double-check all your advice and suggested changes. Maintain a cheerful and helpful tone.";
}

function loadRedlineSetting() {
  const storedSetting = localStorage.getItem("redlineEnabled");
  return storedSetting !== null ? storedSetting === "true" : true; // Default to true (enabled)
}

function saveRedlineSetting(enabled) {
  localStorage.setItem("redlineEnabled", enabled.toString());
}

function loadRedlineAuthor() {
  const storedAuthor = localStorage.getItem("redlineAuthor");
  if (storedAuthor && storedAuthor.trim() !== "") {
    return storedAuthor;
  }
  return DEFAULT_AUTHOR; // Unified default fallback
}

function saveRedlineAuthor(author) {
  if (author !== undefined && author !== null) {
    localStorage.setItem("redlineAuthor", author.toString());
  }
}

function escapeXml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function normalizeLatexForWordMath(latex) {
  return String(latex || "")
    .replace(/```(?:latex|tex)?/gi, "")
    .replace(/```/g, "")
    .replace(/\\begin\{(?:equation|align|aligned|gather|multline)\*?\}/g, "")
    .replace(/\\end\{(?:equation|align|aligned|gather|multline)\*?\}/g, "")
    .replace(/\\\[|\\\]|\$\$/g, "\n")
    .replace(/\\\\/g, "\n")
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean);
}

function mathParagraphOoxml(line) {
  return `<w:p><m:oMathPara><m:oMath>${mathContentOoxml(line)}</m:oMath></m:oMathPara></w:p>`;
}

function mathRunOoxml(text) {
  return `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t xml:space="preserve">${escapeXml(text)}</m:t></m:r>`;
}

function mathContentOoxml(line) {
  const normalized = String(line || "").trim();
  const subscriptMatch = normalized.match(/^([A-Za-z0-9]+)_\{?([A-Za-z0-9]+)\}?$/);
  if (subscriptMatch) {
    return `<m:sSub><m:e>${mathRunOoxml(subscriptMatch[1])}</m:e><m:sub>${mathRunOoxml(subscriptMatch[2])}</m:sub></m:sSub>`;
  }

  const superscriptMatch = normalized.match(/^([A-Za-z0-9]+)\^\{?([A-Za-z0-9]+)\}?$/);
  if (superscriptMatch) {
    return `<m:sSup><m:e>${mathRunOoxml(superscriptMatch[1])}</m:e><m:sup>${mathRunOoxml(superscriptMatch[2])}</m:sup></m:sSup>`;
  }

  return mathRunOoxml(normalized);
}

const GREEK_LATEX_MAP = {
  "\\alpha": "α",
  "\\beta": "β",
  "\\gamma": "γ",
  "\\delta": "δ",
  "\\epsilon": "ε",
  "\\eta": "η",
  "\\theta": "θ",
  "\\lambda": "λ",
  "\\mu": "μ",
  "\\nu": "ν",
  "\\rho": "ρ",
  "\\sigma": "σ",
  "\\tau": "τ",
  "\\phi": "φ",
  "\\omega": "ω",
  "\\Omega": "Ω"
};

function normalizeInlineMathLatex(latex) {
  return String(latex || "")
    .trim()
    .replace(/^\$+|\$+$/g, "")
    .replace(/\\mathrm\{([^}]+)\}/g, "$1")
    .replace(/\s+/g, "");
}

function buildInlineMathSpec(latex, fallbackText = "") {
  let normalized = normalizeInlineMathLatex(latex || fallbackText);
  if (!normalized) return null;

  const fallbackNormalized = normalizeInlineMathLatex(fallbackText);
  if (!normalized.includes("_") && !normalized.includes("^") && /^[A-Z]{2,4}$/.test(normalized)) {
    normalized = `${normalized.charAt(0)}_${normalized.slice(1)}`;
  } else if (normalized === fallbackNormalized && /^[A-Z]{2,4}$/.test(fallbackNormalized)) {
    normalized = `${fallbackNormalized.charAt(0)}_${fallbackNormalized.slice(1)}`;
  }

  const subscriptMatch = normalized.match(/^(.+?)_\{?([^{}]+)\}?$/);
  const superscriptMatch = normalized.match(/^(.+?)\^\{?([^{}]+)\}?$/);
  const baseRaw = subscriptMatch?.[1] || superscriptMatch?.[1] || normalized;
  const scriptRaw = subscriptMatch?.[2] || superscriptMatch?.[2] || "";
  const base = GREEK_LATEX_MAP[baseRaw] || baseRaw.replace(/^\\/, "");
  const script = GREEK_LATEX_MAP[scriptRaw] || scriptRaw.replace(/^\\/, "");
  const text = `${base}${script}`;

  return {
    text,
    italic: true,
    subscriptText: subscriptMatch ? script : "",
    superscriptText: superscriptMatch ? script : ""
  };
}

async function applyInlineMathFormatting(context, range, inlineSpec) {
  if (!inlineSpec) return;
  const insertedRange = range.insertText(inlineSpec.text, Word.InsertLocation.replace);
  insertedRange.font.italic = !!inlineSpec.italic;
  await context.sync();

  if (inlineSpec.subscriptText) {
    await applyNestedScriptFormatting(context, insertedRange, inlineSpec.subscriptText, "subscript");
  }
  if (inlineSpec.superscriptText) {
    await applyNestedScriptFormatting(context, insertedRange, inlineSpec.superscriptText, "superscript");
  }
}

async function replaceTextWithInlineMath(context, targetText, latex, replaceAll = true, options = {}) {
  const textToFind = String(targetText || "").trim();
  const inlineSpec = buildInlineMathSpec(latex || textToFind, textToFind);
  if (!textToFind || !inlineSpec) return 0;

  const searchScope = options.scope === "selection"
    ? context.document.getSelection()
    : context.document.body;
  const ranges = searchScope.search(textToFind, {
    matchCase: options.matchCase !== false,
    matchWholeWord: options.matchWholeWord !== false
  });
  ranges.load("items/text");
  await context.sync();

  const rangesToConvert = replaceAll ? ranges.items : ranges.items.slice(0, 1);
  let convertedCount = 0;
  for (const range of rangesToConvert) {
    await applyInlineMathFormatting(context, range, inlineSpec);
    convertedCount++;
  }
  await context.sync();
  return convertedCount;
}

function buildWordEquationOoxml(latex, title) {
  const parts = [];
  if (title && title.trim()) {
    parts.push(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${escapeXml(title.trim())}</w:t></w:r></w:p>`);
  }

  const lines = normalizeLatexForWordMath(latex);
  if (lines.length === 0) {
    throw new Error("No equation content was provided.");
  }

  for (const line of lines) {
    parts.push(mathParagraphOoxml(line));
  }

  return wrapInDocumentFragment(parts.join(""))
    .replace(
      /<w:document([^>]*)>/,
      '<w:document$1 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    );
}

async function executeInsertWordEquation(latex, location = "cursor", title = "") {
  const checkpointIndex = await createCheckpoint(true);
  const equationOoxml = buildWordEquationOoxml(latex, title);
  const requestedLocation = String(location || "cursor").toLowerCase();

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    if (requestedLocation === "end") {
      context.document.body.insertOoxml(equationOoxml, Word.InsertLocation.end);
    } else if (requestedLocation === "start" || requestedLocation === "beginning" || requestedLocation === "first_page") {
      context.document.body.insertOoxml(equationOoxml, Word.InsertLocation.start);
    } else {
      selection.insertOoxml(equationOoxml, Word.InsertLocation.replace);
    }
    await context.sync();
  });

  return {
    message: requestedLocation === "end"
      ? "Inserted the Word equation at the end."
      : requestedLocation === "start" || requestedLocation === "beginning" || requestedLocation === "first_page"
        ? "Inserted the Word equation at the beginning."
        : "Inserted the Word equation at the cursor.",
    checkpointIndex
  };
}

async function executeConvertTextToWordMath(targetText, latex, scope = "selection", replaceAll = false, targets = null) {
  if (Array.isArray(targets) && targets.length > 0) {
    const checkpointIndex = await createCheckpoint(true);
    let convertedCount = 0;

    await Word.run(async (context) => {
      const trackingState = await setChangeTrackingForAi(context, false, "executeConvertTextToWordMathBatch");
      try {
        for (const target of targets) {
          const textToFind = String(target?.targetText || target?.text || "").trim();
          const equationLatex = String(target?.latex || "").trim();
          if (!textToFind || !equationLatex) continue;

          convertedCount += await replaceTextWithInlineMath(
            context,
            textToFind,
            equationLatex,
            target.replaceAll !== false,
            {
              scope: "document",
              matchCase: target.matchCase !== false,
              matchWholeWord: target.matchWholeWord !== false
            }
          );
        }

        await context.sync();
      } finally {
        await restoreChangeTracking(context, trackingState, "executeConvertTextToWordMathBatch");
      }
    });

    return {
      message: convertedCount > 0
        ? `Converted ${convertedCount} text occurrence(s) to inline math formatting.`
        : "No matching symbols were found to convert to inline math formatting.",
      showToUser: convertedCount > 0,
      checkpointIndex
    };
  }

  const requestedScope = String(scope || "selection").toLowerCase();
  const textToFind = String(targetText || "").trim();
  const equationLatex = String(latex || textToFind || "").trim();

  if (!equationLatex) {
    return {
      message: "No math expression was provided.",
      showToUser: false
    };
  }

  const checkpointIndex = await createCheckpoint(true);
  const inlineSpec = buildInlineMathSpec(equationLatex, textToFind);
  let convertedCount = 0;

  await Word.run(async (context) => {
    const trackingState = await setChangeTrackingForAi(context, false, "executeConvertTextToWordMath");
    try {
      if (requestedScope === "selection" || !textToFind) {
        const selection = context.document.getSelection();
        await applyInlineMathFormatting(context, selection, inlineSpec);
        convertedCount = 1;
      } else {
        convertedCount = await replaceTextWithInlineMath(
          context,
          textToFind,
          equationLatex,
          replaceAll,
          { scope: "document", matchCase: true, matchWholeWord: true }
        );
      }

      await context.sync();
    } finally {
      await restoreChangeTracking(context, trackingState, "executeConvertTextToWordMath");
    }
  });

  return {
    message: convertedCount > 0
      ? `Converted ${convertedCount} text occurrence(s) to inline math formatting.`
      : "No matching text was found to convert to inline math formatting.",
    showToUser: convertedCount > 0,
    checkpointIndex
  };
}

async function executeFormatTextOccurrences(targets, scope = "document") {
  if (!Array.isArray(targets) || targets.length === 0) {
    return {
      message: "No formatting targets were provided.",
      showToUser: false
    };
  }

  const checkpointIndex = await createCheckpoint(true);
  const requestedScope = String(scope || "document").toLowerCase();
  let formattedCount = 0;
  let matchedCount = 0;

  await Word.run(async (context) => {
    const trackingState = await setChangeTrackingForAi(context, loadRedlineSetting(), "executeFormatTextOccurrences");
    try {
      const searchScope = requestedScope === "selection"
        ? context.document.getSelection()
        : context.document.body;

      for (const target of targets) {
        const text = String(target?.text || "").trim();
        if (!text) continue;

        const ranges = searchScope.search(text, {
          matchCase: target.matchCase !== false,
          matchWholeWord: !!target.matchWholeWord
        });
        ranges.load("items/text");
        await context.sync();

        for (const range of ranges.items) {
          matchedCount++;
          let rangeChanged = false;
          if (target.bold !== undefined) range.font.bold = !!target.bold;
          if (target.italic !== undefined) range.font.italic = !!target.italic;
          if (target.underline !== undefined) {
            range.font.underline = target.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
          }
          if (target.bold !== undefined || target.italic !== undefined || target.underline !== undefined) {
            rangeChanged = true;
          }
          if (target.strikethrough !== undefined) {
            range.font.strikeThrough = !!target.strikethrough;
            rangeChanged = true;
          }

          if (target.subscript === true) {
            range.font.subscript = true;
            range.font.superscript = false;
            rangeChanged = true;
          }
          if (target.superscript === true) {
            range.font.superscript = true;
            range.font.subscript = false;
            rangeChanged = true;
          }

          const nestedSubscriptCount = await applyNestedScriptFormatting(context, range, target.subscriptText, "subscript");
          const nestedSuperscriptCount = await applyNestedScriptFormatting(context, range, target.superscriptText, "superscript");
          if (nestedSubscriptCount > 0 || nestedSuperscriptCount > 0) {
            rangeChanged = true;
          }
          if (rangeChanged) {
            formattedCount++;
          }
        }
      }

      await context.sync();
    } finally {
      await restoreChangeTracking(context, trackingState, "executeFormatTextOccurrences");
    }
  });

  return {
    message: formattedCount > 0
      ? `Formatted ${formattedCount} text occurrence(s).`
      : matchedCount > 0
        ? `Found ${matchedCount} matching text occurrence(s), but no nested formatting was applied.`
        : "No matching text was found to format.",
    showToUser: formattedCount > 0,
    checkpointIndex
  };
}

function normalizeWordInsertLocation(location, fallback = Word.InsertLocation.replace) {
  const normalized = String(location || "").toLowerCase();
  if (normalized === "start" || normalized === "before") return Word.InsertLocation.start;
  if (normalized === "end" || normalized === "after") return Word.InsertLocation.end;
  if (normalized === "replace") return Word.InsertLocation.replace;
  return fallback;
}

function applyFontPlan(range, fontPlan = {}) {
  if (fontPlan.bold !== undefined) range.font.bold = !!fontPlan.bold;
  if (fontPlan.italic !== undefined) range.font.italic = !!fontPlan.italic;
  if (fontPlan.underline !== undefined) {
    range.font.underline = fontPlan.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
  }
  if (fontPlan.strikethrough !== undefined) range.font.strikeThrough = !!fontPlan.strikethrough;
  if (fontPlan.subscript !== undefined) {
    range.font.subscript = !!fontPlan.subscript;
    if (fontPlan.subscript) range.font.superscript = false;
  }
  if (fontPlan.superscript !== undefined) {
    range.font.superscript = !!fontPlan.superscript;
    if (fontPlan.superscript) range.font.subscript = false;
  }
  if (fontPlan.color) range.font.color = fontPlan.color;
  if (fontPlan.size) range.font.size = Number(fontPlan.size);
  if (fontPlan.name) range.font.name = String(fontPlan.name);
}

function applyParagraphPlan(paragraph, paragraphPlan = {}) {
  if (paragraphPlan.style) paragraph.style = String(paragraphPlan.style);
  if (paragraphPlan.alignment) paragraph.alignment = paragraphPlan.alignment;
  if (paragraphPlan.spaceBefore !== undefined) paragraph.spaceBefore = Number(paragraphPlan.spaceBefore);
  if (paragraphPlan.spaceAfter !== undefined) paragraph.spaceAfter = Number(paragraphPlan.spaceAfter);
  if (paragraphPlan.lineSpacing !== undefined) paragraph.lineSpacing = Number(paragraphPlan.lineSpacing);
  if (paragraphPlan.leftIndent !== undefined) paragraph.leftIndent = Number(paragraphPlan.leftIndent);
  if (paragraphPlan.firstLineIndent !== undefined) paragraph.firstLineIndent = Number(paragraphPlan.firstLineIndent);
}

async function executeRunWordScript(operations = [], description = "", javascript = "") {
  const script = String(javascript || "").trim();
  if ((!Array.isArray(operations) || operations.length === 0) && !script) {
    return {
      message: "No Word script operations were provided.",
      showToUser: false,
      success: false
    };
  }

  const checkpointIndex = await createCheckpoint(true);
  let appliedCount = 0;
  const errors = [];

  await Word.run(async (context) => {
    const trackingState = await setChangeTrackingForAi(context, loadRedlineSetting(), "executeRunWordScript");
    try {
      if (script) {
        const runGeneratedScript = new Function(
          "Word",
          "Office",
          "context",
          "helpers",
          `"use strict"; return (async () => {\n${script}\n})();`
        );
        await runGeneratedScript(Word, Office, context, {
          buildWordEquationOoxml,
          buildInlineMathSpec,
          applyInlineMathFormatting,
          replaceTextWithInlineMath,
          escapeXml,
          wrapInDocumentFragment
        });
        appliedCount++;
      }

      for (const operation of operations.slice(0, 50)) {
        try {
          const action = String(operation?.action || "").toLowerCase();
          const scope = String(operation?.scope || "selection").toLowerCase();
          const targetText = String(operation?.targetText || operation?.text || "");
          const replaceAll = operation?.replaceAll === true || scope === "document";
          const matchWholeWord = operation?.matchWholeWord !== false;
          const matchCase = operation?.matchCase !== false;

          const getRanges = async () => {
            if (scope === "selection" && !targetText) {
              return [context.document.getSelection()];
            }
            if (!targetText) {
              throw new Error(`${action || "Word script"} requires targetText unless scope is selection.`);
            }
            const searchScope = scope === "selection"
              ? context.document.getSelection()
              : context.document.body;
            const results = searchScope.search(targetText, { matchCase, matchWholeWord });
            results.load("items/text");
            await context.sync();
            return replaceAll ? results.items : results.items.slice(0, 1);
          };

          if (action === "select_text") {
            const ranges = await getRanges();
            if (ranges[0]) {
              ranges[0].select();
              appliedCount++;
            }
          } else if (action === "highlight" || action === "highlight_text") {
            for (const range of await getRanges()) {
              range.font.highlightColor = operation.color || "yellow";
              appliedCount++;
            }
          } else if (action === "format_text") {
            for (const range of await getRanges()) {
              applyFontPlan(range, operation.font || operation);
              appliedCount++;
            }
          } else if (action === "replace_text") {
            for (const range of await getRanges()) {
              const insertedRange = range.insertText(String(operation.replacementText || ""), Word.InsertLocation.replace);
              if (operation.font) applyFontPlan(insertedRange, operation.font);
              appliedCount++;
            }
          } else if (action === "insert_text") {
            const selection = context.document.getSelection();
            const insertedRange = selection.insertText(String(operation.text || ""), normalizeWordInsertLocation(operation.location, Word.InsertLocation.end));
            if (operation.font) applyFontPlan(insertedRange, operation.font);
            appliedCount++;
          } else if (action === "convert_inline_math") {
            const inlineSpec = buildInlineMathSpec(operation.latex || operation.replacementText || targetText, targetText);
            for (const range of await getRanges()) {
              await applyInlineMathFormatting(context, range, inlineSpec);
              appliedCount++;
            }
          } else if (action === "insert_equation" || action === "insert_word_equation") {
            const equationOoxml = buildWordEquationOoxml(operation.latex || "", operation.title || "");
            const selection = context.document.getSelection();
            selection.insertOoxml(equationOoxml, normalizeWordInsertLocation(operation.location, Word.InsertLocation.replace));
            appliedCount++;
          } else if (action === "set_paragraph_format") {
            const paragraphs = scope === "document"
              ? context.document.body.paragraphs
              : context.document.getSelection().paragraphs;
            paragraphs.load("items");
            await context.sync();
            for (const paragraph of paragraphs.items) {
              applyParagraphPlan(paragraph, operation.paragraph || operation);
              appliedCount++;
            }
          } else {
            errors.push(`Unsupported action: ${operation?.action || "(missing)"}`);
          }
        } catch (operationError) {
          errors.push(operationError.message || String(operationError));
        }
      }
      await context.sync();
    } finally {
      await restoreChangeTracking(context, trackingState, "executeRunWordScript");
    }
  });

  const errorSuffix = errors.length > 0 ? ` ${errors.length} operation(s) could not be applied.` : "";
  return {
    message: appliedCount > 0
      ? `Applied ${appliedCount} Word script operation(s).${errorSuffix}`
      : `No Word script operations were applied.${errorSuffix}`,
    showToUser: appliedCount > 0,
    success: appliedCount > 0,
    checkpointIndex,
    description
  };
}

async function applyNestedScriptFormatting(context, parentRange, nestedText, scriptType) {
  const text = normalizeNestedScriptText(nestedText);
  if (!text) return 0;

  const nestedRanges = parentRange.search(text, { matchCase: true, matchWholeWord: false });
  nestedRanges.load("items/text");
  await context.sync();

  let formattedCount = 0;
  for (const nestedRange of nestedRanges.items) {
    if (scriptType === "subscript") {
      nestedRange.font.subscript = true;
      nestedRange.font.superscript = false;
      formattedCount++;
    } else if (scriptType === "superscript") {
      nestedRange.font.superscript = true;
      nestedRange.font.subscript = false;
      formattedCount++;
    }
  }
  return formattedCount;
}

function normalizeNestedScriptText(nestedText) {
  let text = String(nestedText || "").trim();
  if (!text) return "";
  text = text.replace(/^`+|`+$/g, "").trim();
  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
    text = text.slice(1, -1).trim();
  }
  return text;
}

async function setChangeTrackingForAi(context, redlineEnabled, sourceLabel = "AI") {
  let originalMode = null;
  let changed = false;
  let available = false;

  try {
    const doc = context.document;
    doc.load("changeTrackingMode");
    await context.sync();

    available = true;
    originalMode = doc.changeTrackingMode;
    const desiredMode = redlineEnabled ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;

    if (originalMode !== desiredMode) {
      doc.changeTrackingMode = desiredMode;
      await context.sync();
      changed = true;
    }
  } catch (error) {
    console.warn(`[ChangeTracking] ${sourceLabel}: unavailable`, error);
  }

  return { available, originalMode, changed };
}

async function restoreChangeTracking(context, trackingState, sourceLabel = "AI") {
  if (!trackingState || !trackingState.available || !trackingState.changed || trackingState.originalMode === null) {
    return;
  }

  try {
    context.document.changeTrackingMode = trackingState.originalMode;
    await context.sync();
  } catch (error) {
    console.warn(`[ChangeTracking] ${sourceLabel}: restore failed`, error);
  }
}

initAgenticTools({
  loadApiKey,
  loadModel,
  loadSystemMessage,
  loadRedlineSetting,
  loadRedlineAuthor,
  setChangeTrackingForAi,
  restoreChangeTracking,
  SEARCH_LIMITS,
  SAFETY_SETTINGS_BLOCK_NONE,
  API_LIMITS
});

/**
 * Fetches the document's author from Word properties.
 */
async function fetchDocumentAuthor() {
  try {
    let author = "";
    await Word.run(async (context) => {
      const properties = context.document.properties;
      properties.load("lastAuthor, author");
      await context.sync();

      // Use lastAuthor if available, otherwise author
      author = properties.lastAuthor || properties.author || "";
    });
    return author;
  } catch (error) {
    console.warn("Could not fetch document author:", error);
    return "";
  }
}

function loadGlanceSettings() {
  const stored = localStorage.getItem("glanceSettings");
  if (stored) {
    try {
      return JSON.parse(stored);
    } catch (e) {
      console.error("Error parsing glance settings", e);
    }
  }
  // Default fallback
  return [
    { id: 'q1', title: 'Grammar & Spelling', question: 'Are there any glaring spelling or grammatical issues?' },
    { id: 'q2', title: 'Factual Accuracy', question: 'Is this document factually accurate?' }
  ];
}

function saveGlanceSettings(settings) {
  localStorage.setItem("glanceSettings", JSON.stringify(settings));
}

function loadGlanceCollapsedState() {
  return localStorage.getItem(GLANCE_COLLAPSED_STORAGE_KEY) === "true";
}

function saveGlanceCollapsedState(isCollapsed) {
  localStorage.setItem(GLANCE_COLLAPSED_STORAGE_KEY, isCollapsed.toString());
}

function applyGlanceCollapsedState(isCollapsed = loadGlanceCollapsedState()) {
  const container = document.getElementById("glance-container");
  const toggleButton = document.getElementById("toggle-glance-button");
  if (!container || !toggleButton) return;

  container.classList.toggle("collapsed", isCollapsed);
  toggleButton.setAttribute("aria-expanded", (!isCollapsed).toString());
  toggleButton.setAttribute("title", isCollapsed ? "Show Glance results" : "Hide Glance results");
}

function setupAccordion(headerId, contentId) {
  const header = document.getElementById(headerId);
  const content = document.getElementById(contentId);

  if (header && content) {
    header.onclick = () => {
      const isOpen = content.classList.contains("open");

      if (isOpen) {
        content.classList.remove("open");
        header.classList.remove("active");
        // Wait for transition then hide (optional, but keep display block for anim)
        // We rely on max-height: 0 hiding it
      } else {
        content.classList.add("open");
        header.classList.add("active");
      }
    };
  }
}


function renderGlanceMain() {
  const list = document.getElementById("glance-list");
  const container = document.getElementById("glance-container");
  list.innerHTML = "";
  const settings = loadGlanceSettings();

  if (settings.length === 0) {
    if (container) container.style.display = "none";
    return;
  }

  if (container) container.style.display = "block";
  applyGlanceCollapsedState();

  settings.forEach(item => {
    const div = document.createElement("div");
    div.className = "glance-item";
    div.id = `glance-item-${item.id}`;
    div.innerHTML = `
      <div class="glance-header">
        <span id="glance-indicator-${item.id}" class="glance-indicator gray"></span>
        <span class="glance-title">${item.title}</span>
      </div>
      <p id="glance-summary-${item.id}" class="glance-summary">Waiting for analysis...</p>
    `;
    list.appendChild(div);
  });
}

function renderGlanceSettings() {
  const list = document.getElementById("glance-settings-list");
  list.innerHTML = "";
  const settings = loadGlanceSettings();

  settings.forEach((item, index) => {
    const card = document.createElement("div");
    card.className = "glance-settings-card";
    card.dataset.index = index;
    card.dataset.id = item.id;

    // Slimmer layout: Drag handle on left, inputs stacked but compact
    card.innerHTML = `
      <div class="glance-card-header-row">
        <input type="text" class="ms-TextField-field glance-title-input" value="${item.title}" placeholder="Title">
        <span class="drag-handle" title="Drag to reorder">☰</span>
        <button class="delete-card-btn" title="Delete">✕</button>
      </div>
      <textarea class="ms-TextField-field glance-question-input" placeholder="Question (e.g. Is the grammar correct?)" rows="2">${item.question}</textarea>
    `;

    // Event Listeners
    card.querySelector(".delete-card-btn").onclick = () => {
      settings.splice(index, 1);
      saveGlanceSettings(settings);
      renderGlanceSettings();
    };

    const titleInput = card.querySelector(".glance-title-input");
    titleInput.onchange = (e) => {
      settings[index].title = e.target.value;
      saveGlanceSettings(settings);
    };

    const questionInput = card.querySelector(".glance-question-input");
    questionInput.onchange = (e) => {
      settings[index].question = e.target.value;
      saveGlanceSettings(settings);
    };

    // Drag Events - Attach start/end to HANDLE only
    const handle = card.querySelector('.drag-handle');
    handle.draggable = true;
    handle.addEventListener('dragstart', handleDragStart);
    handle.addEventListener('dragend', handleDragEnd);

    // Drop targets are still the CARDS
    card.addEventListener('dragover', handleDragOver);
    card.addEventListener('drop', handleDrop);
    card.addEventListener('dragenter', handleDragEnter);
    card.addEventListener('dragleave', handleDragLeave);

    list.appendChild(card);
  });
}

// Drag and Drop Handlers
let dragSrcEl = null;

function handleDragStart(e) {
  const card = this.closest('.glance-settings-card');
  card.style.opacity = '0.4';
  dragSrcEl = card;
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/html', card.innerHTML);
}

function handleDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  return false;
}

function handleDragToggleClass(e, addClass) {
  const card = e.target.closest('.glance-settings-card');
  if (card) {
    card.classList.toggle('over', addClass);
  }
}

function handleDragEnter(e) {
  handleDragToggleClass(e, true);
}

function handleDragLeave(e) {
  handleDragToggleClass(e, false);
}

function handleDrop(e) {
  e.stopPropagation();

  const targetCard = e.target.closest('.glance-settings-card');

  if (dragSrcEl !== targetCard && targetCard) {
    const list = document.getElementById("glance-settings-list");
    const items = Array.from(list.children);
    const srcIndex = items.indexOf(dragSrcEl);
    const destIndex = items.indexOf(targetCard);

    const settings = loadGlanceSettings();
    const [movedItem] = settings.splice(srcIndex, 1);
    settings.splice(destIndex, 0, movedItem);

    saveGlanceSettings(settings);
    renderGlanceSettings();
  }
  return false;
}

function handleDragEnd(e) {
  const card = this.closest('.glance-settings-card');
  if (card) card.style.opacity = '1';

  const items = document.querySelectorAll('.glance-settings-card');
  items.forEach(function (item) {
    item.classList.remove('over');
  });
}

async function runGlanceChecks() {
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) return;

  const settings = loadGlanceSettings();
  if (settings.length === 0) return;

  // Update UI to showing loading
  settings.forEach(item => {
    const indicator = document.getElementById(`glance-indicator-${item.id}`);
    const summary = document.getElementById(`glance-summary-${item.id}`);
    if (indicator) indicator.className = "glance-indicator gray";
    if (summary) summary.innerText = "Checking...";
  });

  try {
    let docText = "";
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      docText = body.text;
    });

    const model = loadModel('fast'); // Use fast model for glance checks
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${geminiApiKey}`;

    // Prepare prompt for dynamic checks
    let questionsPrompt = "";
    settings.forEach((item, index) => {
      questionsPrompt += `Question ${index + 1} (ID: "${item.id}"): ${item.question}\n`;
    });

    const prompt = `
      Analyze the following document text and answer the following questions.
      Return the result as a JSON object where keys are the Question IDs (e.g., "q1", "q2").
      For each question, provide:
      - "status": "green" (no issues/good), "yellow" (minor issues/caution), or "red" (major issues/bad).
      - "summary": A very brief summary (max 10 words).

      IMPORTANT: Return ONLY the JSON object. Do not include any markdown formatting (like \`\`\`json), conversational text, or explanations.

      Questions:
      ${questionsPrompt}

      Document Text:
      """${docText}""" 
    `;

    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      tools: [{ google_search: {} }],
      safetySettings: SAFETY_SETTINGS_BLOCK_NONE
    };

    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    const result = await response.json();
    const candidate = result.candidates[0];
    let text = candidate.content.parts[0].text;

    // Robust JSON Extraction: Find the first '{' and the last '}'
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      text = jsonMatch[0];
    } else {
      // Fallback cleanup if regex fails (though regex is preferred)
      text = text.replace(/^```json\s*/, "").replace(/^```\s*/, "").replace(/```$/, "").trim();
    }

    const json = JSON.parse(text);

    // Update UI
    settings.forEach(item => {
      const res = json[item.id];
      if (res) {
        const indicator = document.getElementById(`glance-indicator-${item.id}`);
        const summary = document.getElementById(`glance-summary-${item.id}`);
        if (indicator) {
          indicator.className = `glance-indicator ${res.status}`;
          // Add pulse animation
          indicator.classList.add("pulse");
          setTimeout(() => indicator.classList.remove("pulse"), 500);
        }
        if (summary) summary.innerText = res.summary;
      }
    });


  } catch (error) {
    console.error("Glance check failed:", error);
    settings.forEach(item => {
      const summary = document.getElementById(`glance-summary-${item.id}`);
      if (summary) summary.innerText = "Error running check.";
    });
  }
}

// --- Checkpoint Management ---

function getCheckpoints() {
  const checkpointsJson = localStorage.getItem("docCheckpoints");
  return checkpointsJson ? JSON.parse(checkpointsJson) : [];
}

function saveCheckpoints(checkpoints) {
  const MAX_RETRIES = 10; // Maximum number of retry attempts

  let retries = 0;
  while (retries < MAX_RETRIES) {
    try {
      localStorage.setItem("docCheckpoints", JSON.stringify(checkpoints));
      return true; // Success
    } catch (error) {
      if (error.name === 'QuotaExceededError' && checkpoints.length > 1) {
        // Remove 50% of checkpoints (more aggressive pruning)
        const toRemove = Math.max(1, Math.floor(checkpoints.length / 2));
        checkpoints.splice(0, toRemove);
        console.warn(`QuotaExceededError: Removed ${toRemove} oldest checkpoint(s), ${checkpoints.length} remaining. Retrying...`);
        retries++;
      } else if (error.name === 'QuotaExceededError' && checkpoints.length <= 1) {
        // Can't prune anymore, clear all and give up gracefully
        console.warn("Storage quota exceeded. Clearing all checkpoints.");
        try {
          localStorage.removeItem("docCheckpoints");
        } catch (e) { /* ignore */ }
        return false; // Silently fail rather than throw
      } else {
        // Not a quota error
        console.error("Failed to save checkpoints:", error);
        return false; // Silently fail rather than throw
      }
    }
  }

  // If we've exhausted retries, fail gracefully
  console.warn("Unable to save checkpoint after max retries. Clearing checkpoints.");
  try {
    localStorage.removeItem("docCheckpoints");
  } catch (e) { /* ignore */ }
  return false;
}

// function updateCheckpointStatus() { ... } removed as UI is gone.

async function createCheckpoint(silent = false) {
  if (!silent) {
    addMessageToChat("System", "Saving checkpoint...");
  }
  try {
    return await Word.run(async (context) => {
      const ooxml = context.document.body.getOoxml();
      await context.sync();

      // 'ooxml.value' is a base64 string of the entire document body
      const ooxmlLength = ooxml.value.length;
      console.log(`Checkpoint OOXML length: ${ooxmlLength}`);

      const checkpoints = getCheckpoints();

      // Check for quota issues roughly (5MB limit usually)
      let totalSize = 0;
      checkpoints.forEach(c => totalSize += c.length);
      console.log(`Current total checkpoints size: ${totalSize}`);

      let prunedCount = 0;

      // Prune at least MIN_PRUNE_COUNT checkpoints if we need to prune any, to create a buffer
      while ((totalSize + ooxmlLength > STORAGE_LIMITS.SAFE_LIMIT || (prunedCount > 0 && prunedCount < STORAGE_LIMITS.MIN_PRUNE_COUNT)) && checkpoints.length > 0) {
        const removed = checkpoints.shift(); // Remove oldest
        totalSize -= removed.length;
        prunedCount++;
      }

      if (prunedCount > 0) {
        console.warn(`LocalStorage quota exceeded. Removed ${prunedCount} oldest checkpoint(s).`);
        if (!silent) {
          addMessageToChat("System", `Storage full. Removed ${prunedCount} old checkpoint(s) to make space.`);
        }
      }

      checkpoints.push(ooxml.value);
      saveCheckpoints(checkpoints);

      if (!silent) {
        addMessageToChat("System", `Checkpoint saved. Total: ${checkpoints.length}`);
      }

      // Return the index of the newly created checkpoint (0-based)
      return checkpoints.length - 1;
    });
  } catch (error) {
    console.error("Error saving checkpoint:", error);
    if (!silent) {
      addMessageToChat("Error", `Could not save checkpoint. ${error.message}`);
    }
    return -1;
  }
}


async function restoreCheckpoint(index) {
  const checkpoints = getCheckpoints();
  if (index < 0 || index >= checkpoints.length) {
    addMessageToChat("Error", "Invalid checkpoint index.");
    return;
  }

  const msgElement = addMessageToChat("System", `Reverting to checkpoint #${index + 1}...`);

  const targetCheckpointOoxml = checkpoints[index];

  try {
    await Word.run(async (context) => {
      // Disable Track Changes to avoid "Delete All + Insert All" redlines
      const doc = context.document;
      doc.load("changeTrackingMode");
      await context.sync();

      const originalMode = doc.changeTrackingMode;
      if (originalMode !== Word.ChangeTrackingMode.off) {
        doc.changeTrackingMode = Word.ChangeTrackingMode.off;
        await context.sync();
      }

      context.document.body.clear(); // Clear the current document body
      context.document.body.insertOoxml(targetCheckpointOoxml, "Replace");
      await context.sync();

      // Optionally restore track changes, but reverting usually implies going back to a state.
      // If we restore it, we might want to do it cleanly.
      if (originalMode !== Word.ChangeTrackingMode.off) {
        doc.changeTrackingMode = originalMode;
        await context.sync();
      }

      updateSystemMessage(msgElement, "Reverted successfully.");
    });
  } catch (error) {
    console.error("Error reverting checkpoint:", error);
    updateSystemMessage(msgElement, "Error: Could not revert checkpoint.");
  }
}

registerChatUiHandlers({
  onCancelRequest: () => {
    if (currentRequestController) {
      currentRequestController.abort();
      console.log('User cancelled request');
    }
  },
  onRestoreCheckpoint: restoreCheckpoint
});

// --- Chat Feature ---

async function sendChatMessage(modelType = 'fast', messageOverride = null) {
  const chatInput = document.getElementById("chat-input");
  const sendButton = document.getElementById("send-button");
  const thinkButton = document.getElementById("think-button");
  const userMessage = messageOverride || chatInput.value;

  if (userMessage.trim() === "") {
    shakeInput();
    return;
  }

  // Hide any existing retry buttons since conversation is continuing
  hideAllRetryButtons();

  // Reset tool execution tracker for this request
  toolsExecutedInCurrentRequest = [];

  // Sanitize history to remove any hanging function calls from interrupted sessions
  chatHistory = sanitizeHistory(chatHistory);

  // Set up abort controller for this request (allows user cancellation)
  currentRequestController = new AbortController();
  const requestStartTime = Date.now();

  // Lock UI
  chatInput.disabled = true;
  sendButton.disabled = true;
  if (thinkButton) thinkButton.disabled = true;

  // Display user message
  addMessageToChat("User", userMessage);
  chatInput.value = "";

  // Show loading indicator with typing dots and cancel button (yellow for slow, teal for fast)
  const dotColor = modelType === 'slow' ? 'yellow' : 'teal';
  const loadingMsg = createTypingIndicator(dotColor, true); // true = include cancel button
  const chatMessages = document.getElementById("chat-messages");
  chatMessages.appendChild(loadingMsg);
  chatMessages.scrollTop = chatMessages.scrollHeight;




  try {
    // --- Get Document Context ---
    let docText = "";
    let docComments = [];
    let docRedlines = [];
    let docSelection = "";

    await Word.run(async (context) => {
      const body = context.document.body;

      // --- STAGE 1: Critical Text Retrieval ---
      // Fetch current selection & basic text first
      const selection = context.document.getSelection();
      selection.load("text");

      // We'll try enhanced extraction first as it's the gold standard
      try {
        const enhancedContext = await extractEnhancedDocumentContext(context);
        docText = enhancedContext.formattedText;
        console.log(`Enhanced context extracted: ${enhancedContext.paragraphs.length} paragraphs`);
      } catch (enhancedError) {
        console.warn("Enhanced context failed, falling back to simple text", enhancedError);
        // Fallback
        body.load("text");
        await context.sync();
        docText = body.text;
      }

      docSelection = selection.text;

      // Sync to ensure we captured text/selection before trying risky features
      await context.sync();

      // --- STAGE 2: Optional Rich Data (Comments/Redlines) ---
      // These are prone to failure in older Word versions or specific environments
      try {
        const isWordApi14 = Office.context.requirements.isSetSupported("WordApi", "1.4");
        const isWordApi16 = Office.context.requirements.isSetSupported("WordApi", "1.6");

        if (isWordApi14) {
          const comments = context.document.comments;
          comments.load("items/content, items/authorName, items/creationDate");

          let trackedChanges = null;
          if (isWordApi16) {
            try {
              trackedChanges = body.getTrackedChanges();
              trackedChanges.load("items/type, items/text, items/author, items/date");
            } catch (e) { console.warn("Tracked changes not supported (API available but failed)", e); }
          } else {
            console.log("Tracked changes not supported (WordApi 1.6 required)");
          }

          await context.sync(); // syncing specifically for comments/redlines

          // Process optional data
          if (comments && comments.items) {
            docComments = comments.items.map(c => `[Comment by ${c.authorName} on ${c.creationDate}]: ${c.content}`);
          }
          if (trackedChanges && trackedChanges.items) {
            docRedlines = trackedChanges.items.map(tc => `[${tc.type} by ${tc.author} on ${tc.date}]: "${tc.text}"`);
          }
        } else {
          console.log("Optional rich data (comments/redlines) not supported (WordApi 1.4 required)");
        }

      } catch (optionalDataError) {
        if (optionalDataError.name === "RichApi.Error" && optionalDataError.code === "ApiNotFound") {
          console.warn("Could not fetch comments or redlines (API not found despite support check), proceeding with text only.");
        } else {
          console.warn("Could not fetch comments or redlines (API error), proceeding with text only:", optionalDataError);
        }
      }

    });
    // --- Check Document Size ---
    const wordCount = docText.split(/\s+/).length;
    const estimatedTokens = Math.ceil(wordCount * DOCUMENT_LIMITS.TOKEN_MULTIPLIER);

    if (wordCount > DOCUMENT_LIMITS.MAX_WORDS) {
      removeMessage(loadingMsg);
      addMessageToChat("System", `Document is too large to process (approx. ${estimatedTokens} tokens). Please reduce the document size or select a smaller section.`);

      // Re-enable UI
      chatInput.disabled = false;
      sendButton.disabled = false;
      if (thinkButton) thinkButton.disabled = false;

      return;
    }

    // --- Call Gemini API ---
    const geminiApiKey = loadApiKey();
    if (!geminiApiKey) {
      removeMessage(loadingMsg);
      addMessageToChat("Error", "Please set your Gemini API key in the Settings (click the \u2699 icon in the top right).");
      return;
    }

    const geminiModel = loadModel(modelType);
    const useLiveApi = isLiveModel(geminiModel);

    let contextString = "";
    if (docSelection && docSelection.trim() !== "") {
      contextString += `User Highlighted Text:\n"""${docSelection}"""\n\n`;
    }
    if (docText) {
      contextString += `Context from the current document:\n"""${docText}"""\n\n`;
    }
    if (docComments.length > 0) {
      contextString += `Comments in the document:\n${docComments.join("\n")}\n\n`;
    }
    if (docRedlines.length > 0) {
      contextString += `Tracked Changes (Redlines) in the document:\n${docRedlines.join("\n")}\n\n`;
    }

    const prompt = contextString
      ? `${contextString}User Question:\n${userMessage}`
      : userMessage;

    // Add to history
    chatHistory.push({ role: "user", parts: [{ text: prompt }] });

    // Maintain rolling window - but ensure we don't break function call/response pairs
    if (chatHistory.length > 10) {
      chatHistory = maintainHistoryWindow(chatHistory, 10);
    }

    // Define tools
    const tools = [
      {
        function_declarations: [
          {
            name: "apply_redlines",
            description: "Applies suggested edits to the document. Use this tool whenever the user asks to 'edit text', 'change text', 'modify', 'add', 'delete', 'reword', 'rephrase', 'update', 'bold', 'italicize', 'underline', 'strikethrough', or apply inline TEXT FORMATTING to existing paragraphs.\n\nIMPORTANT - FORMATTING RULES:\n- Bold: **text**\n- Italic: *text*\n- Underline: ++text++\n- Strikethrough: ~~text~~\n\nIMPORTANT - LIST RULES:\n- Use Markdown syntax for lists. \n- For Bullet Lists: Use '* item'. For nested items, indent with 4 spaces (e.g., '    * sub-item').\n- For Numbered Lists: Use '1. item', 'a. item', 'i. item', etc. explicitly. \n- For Nested Numbering: Use '1.1.', '1.1.1.' styles if appropriate. \n- DO NOT use simple hyphens ('-') if you intend to create a structured or numbered list. \n- INDENTATION is critical for sub-levels. Use 2 or 4 spaces.\n\nFor full list structure conversions (like turning multiple lines into A., B., C. or 1., 2., 3. list items), prefer the dedicated list tools.\n\nDo NOT suggest changes in the chat; always use this tool to apply them directly. The edits will be applied under track changes (redlines). NEVER say you have applied edits unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The specific instruction for how to edit the document (e.g., 'Change Lessor to Landlord', 'Fix spelling', 'Reword the introduction').",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "insert_comment",
            description: "Inserts comments into the document based on the user's instruction. Use this tool to flag risks, add notes, or review specific sections. NEVER say you have inserted comments unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The instruction for what to comment on and what to say (e.g., 'Flag all risky clauses', 'Comment on the first paragraph').",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "highlight_text",
            description: "Highlights text with a colored background marker (like a highlighter pen). ONLY use this tool when the user EXPLICITLY asks to 'highlight' text. Do NOT use this for formatting requests like 'bold', 'italicize', or general emphasis - those should use apply_redlines with markdown syntax instead. Use this tool ONLY for explicit highlight requests like 'highlight all dates in yellow' or 'mark these terms with highlighting'. NEVER say you have highlighted text unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The instruction for what to highlight (e.g., 'Highlight all dates', 'Mark placeholders').",
                },
                color: {
                  type: "STRING",
                  enum: ["yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "gray25", "gray50", "black", "white"],
                  description: "Optional: highlight color. Default is 'yellow'. Options include: yellow, green, cyan, magenta, blue, red, and dark variants.",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "perform_research",
            description: "Performs a Google Search to answer questions that require external knowledge, facts, or up-to-date information. Use this when the user asks for information not in the document.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The search query to send to Google.",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "insert_word_equation",
            description: "Insert a Microsoft Word equation object from LaTeX. Use this only after the requested equation and variant are clear. If the user did not specify a location, insert at the cursor. Do not use this before researching equation variants when the user asks for a famous equation.",
            parameters: {
              type: "OBJECT",
              properties: {
                latex: {
                  type: "STRING",
                  description: "The final LaTeX equation content only, with no prose.",
                },
                location: {
                  type: "STRING",
                  enum: ["cursor", "end", "start", "first_page"],
                  description: "Where to insert the equation. Default is cursor.",
                },
                title: {
                  type: "STRING",
                  description: "Optional short title only if the user asks for a title or label.",
                },
              },
              required: ["latex"],
            },
          },
          {
            name: "format_text_occurrences",
            description: "Apply direct Word font formatting to existing text occurrences. Use this when the user's requested result is plain Word text formatting such as italic, bold, underline, strikethrough, subscript, or superscript. Use selection scope for selected text, or document scope for repeated occurrences.",
            parameters: {
              type: "OBJECT",
              properties: {
                targets: {
                  type: "ARRAY",
                  items: {
                    type: "OBJECT",
                    properties: {
                      text: {
                        type: "STRING",
                        description: "Exact existing text to find, for example CL or CD."
                      },
                      bold: { type: "BOOLEAN" },
                      italic: { type: "BOOLEAN" },
                      underline: { type: "BOOLEAN" },
                      strikethrough: { type: "BOOLEAN" },
                      subscript: {
                        type: "BOOLEAN",
                        description: "Set true only when the whole matched text should be subscript."
                      },
                      superscript: {
                        type: "BOOLEAN",
                        description: "Set true only when the whole matched text should be superscript."
                      },
                      subscriptText: {
                        type: "STRING",
                        description: "Substring inside the match to make subscript, for example L in CL."
                      },
                      superscriptText: {
                        type: "STRING",
                        description: "Substring inside the match to make superscript."
                      },
                      matchCase: {
                        type: "BOOLEAN",
                        description: "Default true."
                      },
                      matchWholeWord: {
                        type: "BOOLEAN",
                        description: "Default false, useful for short exact terms when needed."
                      }
                    },
                    required: ["text"]
                  }
                },
                scope: {
                  type: "STRING",
                  enum: ["document", "selection"],
                  description: "Use document unless the user explicitly asks to format only the selection."
                }
              },
              required: ["targets"]
            }
          },
          {
            name: "convert_text_to_word_math",
            description: "Replace existing selected or matched inline text with inline mathematical notation while preserving paragraph flow. Use this when the user's requested result is inline math notation inside existing prose. Provide either a single target or batch targets.",
            parameters: {
              type: "OBJECT",
              properties: {
                targets: {
                  type: "ARRAY",
                  description: "Batch conversions for notation cleanup across the document. Use when the user asks to detect symbols, short forms, or convert all/every notation.",
                  items: {
                    type: "OBJECT",
                    properties: {
                      targetText: {
                        type: "STRING",
                        description: "Existing text to replace, for example CL, CD, Re, alpha, eta, omega-z."
                      },
                      latex: {
                        type: "STRING",
                        description: "Final math expression only, for example C_L, C_D, Re, \\alpha, \\eta, \\omega_z."
                      },
                      replaceAll: {
                        type: "BOOLEAN",
                        description: "Default true for batch notation cleanup."
                      },
                      matchCase: {
                        type: "BOOLEAN",
                        description: "Default true."
                      },
                      matchWholeWord: {
                        type: "BOOLEAN",
                        description: "Default true for scientific symbols."
                      }
                    },
                    required: ["targetText", "latex"]
                  }
                },
                targetText: {
                  type: "STRING",
                  description: "Existing text to replace when scope is document. Leave empty to replace current selection."
                },
                latex: {
                  type: "STRING",
                  description: "Final math expression only, for example C_L, C_D, x^2, or \\alpha."
                },
                scope: {
                  type: "STRING",
                  enum: ["selection", "document"],
                  description: "Use selection when the user says select/selected word/text or here. Use document only when the user asks all/everywhere."
                },
                replaceAll: {
                  type: "BOOLEAN",
                  description: "True only when the user asks all/every occurrence."
                }
              },
              required: ["latex"]
            }
          },
          {
            name: "run_word_script",
            description: "Run Microsoft Word Office.js code or an ordered JSON plan to modify the open document. Use this freely for document-editing tasks when direct scripting is more capable than the narrow tools. A checkpoint is created before execution so the user can revert. If using javascript, write only the body of an async function that receives Word, Office, context, and helpers; call await context.sync() after queued Word operations.",
            parameters: {
              type: "OBJECT",
              properties: {
                description: {
                  type: "STRING",
                  description: "Brief plain-language summary of what the plan will do."
                },
                javascript: {
                  type: "STRING",
                  description: "Office.js script body to run inside Word.run. You may use context.document, Word.InsertLocation, context.sync(), helpers.buildWordEquationOoxml(latex,title), and helpers.replaceTextWithInlineMath(context,targetText,latex,replaceAll,options). For CL to C subscript L across the document, use: const count = await helpers.replaceTextWithInlineMath(context, 'CL', 'C_L', true, { scope: 'document', matchCase: true, matchWholeWord: true }); Do not wrap in function syntax."
                },
                operations: {
                  type: "ARRAY",
                  description: "Optional ordered Word operations if a JSON plan is easier than javascript.",
                  items: {
                    type: "OBJECT",
                    properties: {
                      action: {
                        type: "STRING",
                        enum: [
                          "select_text",
                          "highlight_text",
                          "format_text",
                          "replace_text",
                          "insert_text",
                          "convert_inline_math",
                          "insert_word_equation",
                          "set_paragraph_format"
                        ],
                        description: "Allowed operation to execute."
                      },
                      scope: {
                        type: "STRING",
                        enum: ["selection", "document"],
                        description: "Use selection for selected/current text; use document for all/every/throughout."
                      },
                      targetText: {
                        type: "STRING",
                        description: "Exact existing text to find. Required except for current-selection operations and insertion."
                      },
                      replacementText: {
                        type: "STRING",
                        description: "Replacement text for replace_text or source text for conversion."
                      },
                      text: {
                        type: "STRING",
                        description: "Text to insert or find, depending on action."
                      },
                      latex: {
                        type: "STRING",
                        description: "LaTeX content for math conversion/equation actions, with no prose."
                      },
                      title: {
                        type: "STRING",
                        description: "Optional equation title only when requested."
                      },
                      location: {
                        type: "STRING",
                        enum: ["cursor", "start", "end", "replace", "before", "after"],
                        description: "Insertion location. Default is cursor/replace for equation, end for inserted text."
                      },
                      replaceAll: {
                        type: "BOOLEAN",
                        description: "True for all/every occurrence. False for one occurrence."
                      },
                      matchCase: {
                        type: "BOOLEAN",
                        description: "Default true."
                      },
                      matchWholeWord: {
                        type: "BOOLEAN",
                        description: "Default true."
                      },
                      color: {
                        type: "STRING",
                        description: "Highlight or font color, for example yellow or #ff0000."
                      },
                      font: {
                        type: "OBJECT",
                        properties: {
                          bold: { type: "BOOLEAN" },
                          italic: { type: "BOOLEAN" },
                          underline: { type: "BOOLEAN" },
                          strikethrough: { type: "BOOLEAN" },
                          subscript: { type: "BOOLEAN" },
                          superscript: { type: "BOOLEAN" },
                          color: { type: "STRING" },
                          size: { type: "NUMBER" },
                          name: { type: "STRING" }
                        },
                        description: "Direct font formatting to apply."
                      },
                      paragraph: {
                        type: "OBJECT",
                        properties: {
                          style: { type: "STRING" },
                          alignment: { type: "STRING" },
                          spaceBefore: { type: "NUMBER" },
                          spaceAfter: { type: "NUMBER" },
                          lineSpacing: { type: "NUMBER" },
                          leftIndent: { type: "NUMBER" },
                          firstLineIndent: { type: "NUMBER" }
                        },
                        description: "Paragraph formatting to apply."
                      }
                    },
                    required: ["action"]
                  }
                }
              }
            }
          },
          {
            name: "navigate_to_section",
            description: "Navigates to and selects a specific section of the document. Use this when the user asks to go to, scroll to, find, or jump to a particular part of the document (e.g., 'go to the introduction', 'scroll to paragraph 5', 'find the signature block', 'show me the definitions section'). This helps users quickly locate relevant content without manually scrolling.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The navigation instruction describing what section to go to (e.g., 'go to paragraph 3', 'find the table of contents', 'scroll to the conclusion', 'show me where parties are defined').",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "edit_list",
            description: "Edit an entire list as a unit. Use this when you need to modify, add, or remove items from a bulleted or numbered list. This preserves list formatting and structure better than apply_redlines. Look for paragraphs with |ListNumber or |ListBullet in the context. For numbered lists, you can specify different numbering styles: '1, 2, 3' (decimal - default), 'a, b, c' (lowerAlpha), 'A, B, C' (upperAlpha), 'i, ii, iii' (lowerRoman), or 'I, II, III' (upperRoman). NEVER say you have edited a list unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                startParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the FIRST item in the list (e.g., 3 for [P3])",
                },
                endParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the LAST item in the list",
                },
                newItems: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "The new list items in order. Each string is one list item text (without bullets/numbers).",
                },
                listType: {
                  type: "STRING",
                  enum: ["bullet", "numbered"],
                  description: "The type of list to create",
                },
                numberingStyle: {
                  type: "STRING",
                  enum: ["decimal", "lowerAlpha", "upperAlpha", "lowerRoman", "upperRoman"],
                  description: "Optional: For numbered lists, the numbering style to use. Default is 'decimal' (1, 2, 3). Options: 'decimal' (1, 2, 3), 'lowerAlpha' (a, b, c), 'upperAlpha' (A, B, C), 'lowerRoman' (i, ii, iii), 'upperRoman' (I, II, III).",
                },
              },
              required: ["startParagraphIndex", "endParagraphIndex", "newItems", "listType"],
            },
          },
          {
            name: "insert_list_item",
            description: "Insert a single list item after a specific paragraph. Use this for surgical additions to an existing list - it inherits the numbering format from the paragraph you insert after. Much better than edit_list when you only need to add one or two items. Do NOT include numbering markers in the text - Word will add them automatically.",
            parameters: {
              type: "OBJECT",
              properties: {
                afterParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index to insert after (e.g., 5 to insert after [P5])",
                },
                text: {
                  type: "STRING",
                  description: "The text content of the new list item (WITHOUT any numbering like '1.' or '1.1.' - Word adds these automatically)",
                },
                indentLevel: {
                  type: "INTEGER",
                  description: "Optional: Relative indentation from the paragraph you're inserting after. Allowed values: -1 (one level shallower), 0 (same level, default), 1 (one level deeper). Values outside -1..1 are treated as invalid and clamped.",
                },
              },
              required: ["afterParagraphIndex", "text"],
            },
          },
          {
            name: "edit_table",
            description: "Edit a table as a unit. Use this when you need to modify table content, add/remove rows or columns. This preserves table formatting. Look for paragraphs with |T:row,col in the context. NEVER say you have edited a table unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                paragraphIndex: {
                  type: "INTEGER",
                  description: "Any paragraph index that is part of the table (has T:row,col marker)",
                },
                action: {
                  type: "STRING",
                  enum: ["replace_content", "add_row", "delete_row", "update_cell"],
                  description: "The table operation to perform",
                },
                content: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "For replace_content: 2D array of strings [[row1cells], [row2cells]]. For add_row: array of cell values. For update_cell: single-element array with new text.",
                },
                targetRow: {
                  type: "INTEGER",
                  description: "For add_row/delete_row/update_cell: the 0-based row index",
                },
                targetColumn: {
                  type: "INTEGER",
                  description: "For update_cell: the 0-based column index",
                },
              },
              required: ["paragraphIndex", "action"],
            },
          },
          {
            name: "edit_section",
            description: "Edit a document section as a unit. Use this for legal contracts where numbered/lettered items serve as section headers (marked with §) followed by body text (marked with §N). This preserves the section structure and list numbering. NEVER say you have edited a section unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                sectionHeaderIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the section header (the list item marked with §, e.g., '1. Definitions')",
                },
                newHeaderText: {
                  type: "STRING",
                  description: "Optional: new text for the section header. The list number/letter is automatically preserved.",
                },
                newBodyParagraphs: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "Optional: new body paragraphs for this section. Each string becomes one paragraph. Omit to keep existing body.",
                },
                preserveSubsections: {
                  type: "BOOLEAN",
                  description: "If true, only edits body text until the next subsection. If false or omitted, replaces entire section including subsections.",
                },
              },
              required: ["sectionHeaderIndex"],
            },
          },
          {
            name: "convert_headers_to_list",
            description: "Convert non-contiguous headers to a numbered list. Use this when headers like '1. PURPOSE', '2. DEFINITION', '3. EXCLUSIONS' have body text between them and need to be converted to a proper auto-numbered list. The tool strips manual numbering and creates a Word list where all headers share continuous numbering. Supports different formats: 1,2,3 or a,b,c or i,ii,iii. NEVER say you have converted headers unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                paragraphIndices: {
                  type: "ARRAY",
                  items: { type: "INTEGER" },
                  description: "Array of 1-based paragraph indices of the headers to convert (e.g., [3, 7, 15] for headers at P3, P7, P15)",
                },
                newHeaderTexts: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "Optional: new text for each header (without numbers). If omitted, existing text is used with manual numbers stripped.",
                },
                numberingFormat: {
                  type: "STRING",
                  enum: ["arabic", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman"],
                  description: "Optional: numbering format. 'arabic' = 1,2,3 (default), 'lowerLetter' = a,b,c, 'upperLetter' = A,B,C, 'lowerRoman' = i,ii,iii, 'upperRoman' = I,II,III",
                },
              },
              required: ["paragraphIndices"],
            },
          },
        ],
      },
    ];

    const systemInstruction = {
      parts: [
        {
          text: loadSystemMessage() + `\\n\\nDOCUMENT CONTEXT:
The document text includes internal markers such as [P#], list levels, table positions, and section markers. Use them only to locate content for tool calls. Do not mention these markers to the user.

AGENT BEHAVIOR:
- Follow the user's latest chat instruction first.
- Think like a Word co-worker: inspect the document context, decide the action, and execute it.
- When the user asks for a document change, prefer run_word_script with Office.js if that gives you the most direct control.
- Prefer the current selection when the user says selected text, this, here, or highlighted text.
- Use the whole document when the user says all, every, throughout, or when the request clearly applies globally.
- If you need to edit Word formatting, equations, paragraphs, tables, comments, selection, or styles, write Office.js in run_word_script instead of only explaining.
- For inline notation such as CL -> C_L, CD -> C_D, Re, alpha, omega, use run_word_script javascript with helpers.replaceTextWithInlineMath(context, targetText, latex, replaceAll, options). Example: const count = await helpers.replaceTextWithInlineMath(context, 'CL', 'C_L', true, { scope: 'document', matchCase: true, matchWholeWord: true });
- If the request is genuinely ambiguous, ask one short clarification.
- After a successful tool call, give a full chat response explaining what you did, which tool/script path you used, what changed in the document, and any limitation or uncertainty. Do not expose hidden chain-of-thought, but do explain your practical reasoning and the action result clearly.

AVAILABLE TOOL INTENT:
- apply_redlines: rewrite, replace, delete, add, or edit document text with tracked changes.
- format_text_occurrences: apply normal Word font formatting to existing text.
- convert_text_to_word_math: convert existing inline notation inside prose while preserving paragraph flow.
- insert_word_equation: insert a standalone Word equation from LaTeX.
- insert_comment: add comments.
- highlight_text: highlight text.
- run_word_script: run generated Office.js or a JSON plan against the open Word document with a revert checkpoint.
- edit_list, insert_list_item, convert_headers_to_list: manipulate lists.
- edit_table: manipulate tables.
- edit_section: manipulate structured sections.
- perform_research: look up external information when needed.`,
        },
      ],
    };

    // --- Tool Execution Loop with Multi-Tier Recovery ---
    let loopCount = 0;
    let keepLooping = true;
    let currentRecoveryTier = 0;  // 0=normal, 1=validate pairs, 2=remove all pairs, 3=fresh start, 4=graceful degrade
    const originalUserMessage = prompt;  // Save for Tier 3 recovery
    let consecutiveNoProgressToolLoops = 0;
    let lastNoProgressSignature = "";

    while (keepLooping && loopCount < DOCUMENT_LIMITS.MAX_LOOPS) {
      loopCount++;
      console.log(`Starting chat loop iteration ${loopCount} (recovery tier: ${currentRecoveryTier})`);

      // Check for user cancellation
      if (currentRequestController && currentRequestController.signal.aborted) {
        console.log('Request cancelled by user during loop');
        removeMessage(loadingMsg);
        addMessageToChat("System", "Request cancelled.");
        keepLooping = false;
        break;
      }

      // Check for overall timeout
      const elapsedTime = Date.now() - requestStartTime;
      if (elapsedTime > TIMEOUT_LIMITS.TOTAL_REQUEST_TIMEOUT_MS) {
        console.warn(`Overall request timeout exceeded: ${elapsedTime}ms`);
        removeMessage(loadingMsg);

        // If some tools executed successfully, show partial success
        if (toolsExecutedInCurrentRequest.length > 0) {
          const successMessage = generateSuccessMessage(toolsExecutedInCurrentRequest);
          const timeoutGuidance = "\n\nThe request timed out. Gemini 2.5 Flash is the recommended fast model for document edits.";

          if (successMessage) {
            addMessageToChat("System", successMessage + "\n\n*(Request timed out after completing some changes)*" + timeoutGuidance);
          } else {
            addMessageToChat("Error", "Request timed out. Some changes may have been applied." + timeoutGuidance);
          }
        } else {
          addMessageToChat("Error", "The request timed out before any document changes were applied. Gemini 2.5 Flash is the recommended fast model for document edits.");

          // Discard the timed out request from history to allow user to continue clean
          // Remove the last user message we added for this request
          // (The one pushed at `chatHistory.push({ role: "user", parts: [{ text: prompt }] });`)
          // We only remove it if we haven't successfully done tools that we want to keep context for.
          if (chatHistory.length > 0 && chatHistory[chatHistory.length - 1].role === "user") {
            console.log("Discarding timed out request from history");
            chatHistory.pop();
          }
        }
        keepLooping = false;
        break;
      }

      // Prepare payload with current history
      const payload = {
        contents: chatHistory,
        systemInstruction: systemInstruction,
        tools: tools,
        safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
        generationConfig: {
          maxOutputTokens: API_LIMITS.MAX_OUTPUT_TOKENS
        },
      };

      console.log("Sending Chat History to API:", JSON.stringify(chatHistory, null, 2));

      let result;
      try {
        result = useLiveApi
          ? await callGeminiLiveAsGenerateContent(geminiApiKey, geminiModel, payload)
          : await callGeminiWithModelFallback(geminiApiKey, geminiModel, payload);
      } catch (apiError) {
        console.error(`API Error on iteration ${loopCount}:`, apiError);

        // Check if this is a function call/response mismatch error
        const isFunctionCallError = apiError.message && (
          apiError.message.includes("function response turn comes immediately after a function call turn") ||
          apiError.message.includes("function call turn comes immediately after a user turn or after a function response turn")
        );

        if (isFunctionCallError) {
          currentRecoveryTier++;
          console.warn(`Function call error detected. Escalating to recovery tier ${currentRecoveryTier}`);

          if (currentRecoveryTier === 1) {
            // Tier 1: Validate and clean history pairs
            console.log("Tier 1: Validating history pairs...");
            const originalLength = chatHistory.length;
            chatHistory = validateHistoryPairs(chatHistory);
            console.log(`History cleaned: ${originalLength} -> ${chatHistory.length} messages`);
            loopCount = 0;  // Reset to retry
            continue;
          } else if (currentRecoveryTier === 2) {
            // Tier 2: Remove ALL function pairs
            console.log("Tier 2: Removing all function call/response pairs...");
            chatHistory = removeAllFunctionPairs(chatHistory);
            console.log(`History after removing function pairs: ${chatHistory.length} messages`);
            loopCount = 0;
            continue;
          } else if (currentRecoveryTier === 3) {
            // Tier 3: Fresh start with original context
            console.log("Tier 3: Creating fresh start with original context...");
            chatHistory = createFreshStartWithContext(originalUserMessage);
            console.log(`History reset to fresh start: ${chatHistory.length} messages`);
            loopCount = 0;
            continue;
          } else {
            // Tier 4: Graceful degradation
            console.log("Tier 4: All recovery attempts failed. Checking for graceful degradation...");
            removeMessage(loadingMsg);

            const successMessage = generateSuccessMessage(toolsExecutedInCurrentRequest);
            if (successMessage) {
              addMessageToChat("System", successMessage + "\n\n*(Conversation refreshed)*");
              // Reset history for next request
              chatHistory = [];
            } else {
              addMessageToChat("Error", "I encountered an issue with the conversation. Please try again.");
            }
            keepLooping = false;
            break;
          }
        }

        // Non-recoverable errors after successful tool execution
        if (loopCount > 1 && toolsExecutedInCurrentRequest.length > 0) {
          console.warn("Stopping loop due to API error after successful tool execution.");
          const successMessage = generateSuccessMessage(toolsExecutedInCurrentRequest);
          if (successMessage) {
            if (loadingMsg) {
              updateSystemMessage(loadingMsg, successMessage + "\n\n*(Conversation refreshed)*");
            } else {
              addMessageToChat("System", successMessage + "\n\n*(Conversation refreshed)*");
            }
            chatHistory = [];
          }
          keepLooping = false;
          break;
        } else {
          throw apiError;
        }
      }

      console.log("Gemini chat raw result:", JSON.stringify(result, null, 2));

      if (!result.candidates || !Array.isArray(result.candidates) || result.candidates.length === 0) {
        throw new Error("Gemini response contained no candidates.");
      }

      const candidate = result.candidates[0];
      let parts = [];
      let content = candidate.content;

      if (content && content.parts && Array.isArray(content.parts)) {
        parts = content.parts;
      } else if (
        (candidate.finishReason === "MALFORMED_FUNCTION_CALL" || candidate.finishReason === "UNEXPECTED_TOOL_CALL")
        && (candidate.finishMessage || (candidate.content && candidate.content.parts))
      ) {
        console.warn(`Gemini returned ${candidate.finishReason}. Attempting to recover...`, candidate.finishMessage || candidate.content);

        const toolNames = [
          "apply_redlines",
          "insert_comment",
          "highlight_text",
          "perform_research",
          "format_text_occurrences",
          "convert_text_to_word_math",
          "navigate_to_section",
          "edit_list",
          "insert_list_item",
          "edit_table",
          "edit_section",
          "convert_headers_to_list"
        ];

        const tryParseArgs = (rawArgs) => {
          if (!rawArgs || typeof rawArgs !== "string") return null;
          const trimmed = rawArgs.trim();
          if (!trimmed) return {};

          try {
            return JSON.parse(trimmed);
          } catch (_) {
            // Fall through to tolerant parser.
          }

          try {
            const normalized = trimmed
              .replace(/^\(\s*/, "")
              .replace(/\s*\)\s*$/, "")
              .replace(/,\s*([}\]])/g, "$1")
              .replace(/([{,]\s*)([A-Za-z_][A-Za-z0-9_]*)\s*:/g, '$1"$2":')
              .replace(/'([^'\\]*(?:\\.[^'\\]*)*)'/g, (_, s) => `"${s.replace(/"/g, '\\"')}"`);
            return JSON.parse(normalized);
          } catch (_) {
            return null;
          }
        };

        const parseMalformedEditListArgs = (rawArgs) => {
          if (!rawArgs || typeof rawArgs !== "string") return null;

          const parseIntField = (fieldName) => {
            const match = rawArgs.match(new RegExp(`${fieldName}\\s*:\\s*(\\d+)`, "i"));
            return match ? parseInt(match[1], 10) : null;
          };

          const parseStringField = (fieldName) => {
            const match = rawArgs.match(new RegExp(`${fieldName}\\s*:\\s*([^,}\\]]+)`, "i"));
            return match ? String(match[1]).trim().replace(/^["']|["']$/g, "") : null;
          };

          const startParagraphIndex = parseIntField("startParagraphIndex");
          const endParagraphIndex = parseIntField("endParagraphIndex");
          const listTypeRaw = parseStringField("listType");
          const numberingStyle = parseStringField("numberingStyle");

          if (!startParagraphIndex || !endParagraphIndex) {
            return null;
          }

          const listType = (listTypeRaw || "numbered").toLowerCase();
          const normalizedListType = listType === "bullet" ? "bullet" : "numbered";

          let newItems = [];
          const itemsMatch = rawArgs.match(/newItems\s*:\s*\[([\s\S]*?)\](?=\s*,\s*[A-Za-z_][A-Za-z0-9_]*\s*:|\s*$)/i);
          if (itemsMatch && itemsMatch[1]) {
            const itemsRaw = itemsMatch[1]
              .replace(/\r?\n/g, " ")
              .replace(/\s+/g, " ")
              .trim();

            if (itemsRaw) {
              // Common malformed pattern: unquoted sentence-like items separated by " , "
              const sentenceSplit = itemsRaw
                .split(/(?<=[.;!?])\s*,\s*(?=[A-Z0-9(“"'])/)
                .map((s) => s.trim().replace(/^["']|["']$/g, ""))
                .filter(Boolean);

              if (sentenceSplit.length > 0) {
                newItems = sentenceSplit;
              } else {
                // Fallback for simpler malformed arrays.
                newItems = itemsRaw
                  .split(/\s*,\s*/)
                  .map((s) => s.trim().replace(/^["']|["']$/g, ""))
                  .filter(Boolean);
              }
            }
          }

          const parsed = {
            startParagraphIndex,
            endParagraphIndex,
            listType: normalizedListType
          };

          if (numberingStyle) {
            parsed.numberingStyle = numberingStyle;
          }

          if (newItems.length > 0) {
            parsed.newItems = newItems;
          }

          return parsed;
        };

        const parseMalformedConvertHeadersArgs = (rawArgs) => {
          if (!rawArgs || typeof rawArgs !== "string") return null;
          const indicesMatch = rawArgs.match(/paragraphIndices\s*:\s*\[([^\]]*)\]/i);
          if (!indicesMatch || !indicesMatch[1]) {
            return null;
          }

          const paragraphIndices = indicesMatch[1]
            .split(/\s*,\s*/)
            .map((v) => parseInt(v.trim(), 10))
            .filter((v) => Number.isFinite(v));

          if (paragraphIndices.length === 0) {
            return null;
          }

          const numberingFormatMatch = rawArgs.match(/numberingFormat\s*:\s*([^,}\]]+)/i);
          const numberingFormat = numberingFormatMatch
            ? String(numberingFormatMatch[1]).trim().replace(/^["']|["']$/g, "")
            : undefined;

          return numberingFormat
            ? { paragraphIndices, numberingFormat }
            : { paragraphIndices };
        };

        let recoveredFunctionCall = null;
        for (const toolName of toolNames) {
          const regex = new RegExp(`${toolName}\\s*\\(?\\s*(\\{[\\s\\S]*?\\})\\s*\\)?`, "i");
          const match = candidate.finishMessage.match(regex);
          if (!match || !match[1]) {
            continue;
          }

          let parsedArgs = tryParseArgs(match[1]);
          if ((!parsedArgs || typeof parsedArgs !== "object" || Array.isArray(parsedArgs)) && toolName === "edit_list") {
            parsedArgs = parseMalformedEditListArgs(match[1]);
          }
          if ((!parsedArgs || typeof parsedArgs !== "object" || Array.isArray(parsedArgs)) && toolName === "convert_headers_to_list") {
            parsedArgs = parseMalformedConvertHeadersArgs(match[1]);
          }
          if (!parsedArgs || typeof parsedArgs !== "object" || Array.isArray(parsedArgs)) {
            continue;
          }

          recoveredFunctionCall = {
            name: toolName,
            args: parsedArgs
          };
          break;
        }

        // Legacy fallback: recover redline instruction from raw malformed text.
        if (!recoveredFunctionCall) {
          const redlineMatch = candidate.finishMessage.match(/apply_redlines\s*\{\s*instruction\s*:\s*(.*)\s*\}/s);
          if (redlineMatch && redlineMatch[1]) {
            recoveredFunctionCall = {
              name: "apply_redlines",
              args: { instruction: redlineMatch[1].trim() }
            };
          }
        }

        if (recoveredFunctionCall) {
          console.log("Recovered malformed tool call:", recoveredFunctionCall.name, recoveredFunctionCall.args);
          parts = [{ functionCall: recoveredFunctionCall }];
          // Ensure content has the proper structure with role
          if (!content || !content.role) {
            content = { role: "model", parts: parts };
          } else {
            content.parts = parts;
          }
        }
      }

      if (parts.length === 0) {
        // Handle empty STOP responses gracefully (silent success)
        if (candidate.finishReason === "STOP") {
          console.log("Gemini returned empty parts with finishReason: STOP. Treating as silent success.");
          const successMessage = generateSuccessMessage(toolsExecutedInCurrentRequest);
          parts = [{
            text: successMessage
              ? `${successMessage}\n\nI completed the Word action and preserved the revert checkpoint shown above.`
              : "Task completed successfully."
          }];
        } else if (candidate.finishReason === "UNEXPECTED_TOOL_CALL" || candidate.finishReason === "MALFORMED_FUNCTION_CALL") {
          // The model tried to call a tool but the call was malformed/unexpected
          // and we couldn't recover it. Ask the model to retry without a tool call.
          console.warn(`Gemini returned ${candidate.finishReason} with no recoverable data. Asking model to retry.`);
          chatHistory.push({
            role: "model",
            parts: [{ text: `I encountered an issue trying to use a tool (${candidate.finishReason}). Let me try a different approach.` }]
          });
          chatHistory.push({
            role: "user",
            parts: [{ text: "Your previous tool call was malformed. Please try again — either rephrase the tool call with valid arguments, or respond with text only." }]
          });
          continue;
        } else {
          console.error("Gemini candidate missing content.parts:", candidate);

          let diagnosticInfo = `Finish Reason: ${candidate.finishReason || 'NOT_FOUND'}`;

          // Check for safety ratings that might have triggered an empty response
          if (candidate.safetyRatings && Array.isArray(candidate.safetyRatings)) {
            const highRatings = candidate.safetyRatings.filter(r => r.probability !== "NEGLIGIBLE");
            if (highRatings.length > 0) {
              diagnosticInfo += ` | Safety: ${highRatings.map(r => `${r.category}:${r.probability}`).join(', ')}`;
            }
          }

          // Check for specific finish reasons like SAFETY or RECITATION
          if (candidate.finishReason === "SAFETY") {
            diagnosticInfo += " | Content blocked by safety filters.";
          } else if (candidate.finishReason === "RECITATION") {
            diagnosticInfo += " | Content blocked due to copyright/recitation filters.";
          }

          throw new Error(`Gemini response was missing content.parts. ${diagnosticInfo}`);
        }
      }

      console.log("Gemini chat content.parts:", parts);

      // --- Thought Signature Handling ---
      // Check for thought/reasoning parts to potentially log or handle separately
      const thinkingPart = parts.find(p => p.thought || p.thought_signature || p.thoughtSignature);
      if (thinkingPart) {
        console.log("Model Reasoning detected:", thinkingPart.thought || thinkingPart.thought_signature || thinkingPart.thoughtSignature);
      }

      // Check for ALL function calls in the response
      const functionCallParts = parts.filter((part) => part.functionCall);

      if (functionCallParts.length > 0) {
        // If this is the first loop, remove the "Thinking..." message so we can show tool status
        // Keep loading message visible during tool execution


        // Execute ALL function calls and collect responses
        const functionResponses = [];
        const mutatingToolNames = new Set([
          "apply_redlines",
          "insert_comment",
          "highlight_text",
          "insert_word_equation",
          "format_text_occurrences",
          "convert_text_to_word_math",
          "run_word_script",
          "edit_list",
          "insert_list_item",
          "edit_table",
          "edit_section",
          "convert_headers_to_list"
        ]);
        let attemptedMutatingToolsThisLoop = 0;
        let successfulMutatingToolsThisLoop = 0;
        const failedMutationSignatures = [];

        for (const functionCallPart of functionCallParts) {
          const functionCall = functionCallPart.functionCall;
          const args = functionCall.args;
          const instruction = args.instruction;

          // Update loading message status
          if (loadingMsg) {
            const toolFriendlyNames = {
              "apply_redlines": `Applying edits: "${instruction}"...`,
              "insert_comment": `Inserting comments: "${instruction}"...`,
              "highlight_text": `Highlighting text: "${instruction}"...`,
              "perform_research": `Researching: "${instruction}"...`,
              "insert_word_equation": "Inserting Word equation...",
              "format_text_occurrences": "Applying text formatting...",
              "convert_text_to_word_math": "Converting text to Word math...",
              "run_word_script": "Running Word action plan...",
              "navigate_to_section": `Navigating to: "${instruction}"...`
            };
            const statusText = toolFriendlyNames[functionCall.name] || "Working...";
            updateSystemMessage(loadingMsg, statusText);
          }


          let toolResult = "";
          let toolSucceeded = false;

          if (functionCall.name === "apply_redlines") {
            const checkpointIndex = await createCheckpoint(true);
            const result = await executeRedline(instruction, docText);
            toolResult = result.message;
            toolSucceeded = !!result.showToUser;

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: result.showToUser
            });

            // Only show to user if there were actual changes or a true error
            if (result.showToUser) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              console.log(`Fallback in progress (0 edits): ${toolResult}`);
            }

          } else if (functionCall.name === "insert_comment") {
            const checkpointIndex = await createCheckpoint(true);
            const result = await executeComment(instruction, docText);
            toolResult = result.message;
            toolSucceeded = !!result.showToUser;

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: result.showToUser
            });

            if (result.showToUser) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              console.log(`Fallback in progress (0 comments): ${toolResult}`);
            }

          } else if (functionCall.name === "highlight_text") {
            const checkpointIndex = await createCheckpoint(true);
            const highlightColor = args.color || "yellow";
            const result = await executeHighlight(instruction, docText, highlightColor);
            toolResult = result.message;
            toolSucceeded = !!result.showToUser;

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: result.showToUser
            });

            if (result.showToUser) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              console.log(`Fallback in progress (0 highlights): ${toolResult}`);
            }

          } else if (functionCall.name === "perform_research") {
            updateSystemMessage(loadingMsg, `Researching: "${instruction}"...`);
            toolResult = await executeResearch(instruction);
            toolSucceeded = true;

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: true
            });

            updateSystemMessage(loadingMsg, `Found search results for: "${instruction}"`);
          } else if (functionCall.name === "insert_word_equation") {
            const result = await executeInsertWordEquation(
              args.latex,
              args.location || "cursor",
              args.title || ""
            );
            toolResult = result.message;
            toolSucceeded = true;

            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: "insert_word_equation",
              result: toolResult,
              success: true
            });

            updateSystemMessage(loadingMsg, toolResult, result.checkpointIndex);
          } else if (functionCall.name === "format_text_occurrences") {
            const result = await executeFormatTextOccurrences(
              args.targets || [],
              args.scope || "document"
            );
            toolResult = result.message;
            toolSucceeded = !!result.showToUser;

            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: "format_text_occurrences",
              result: toolResult,
              success: result.showToUser
            });

            updateSystemMessage(loadingMsg, toolResult, result.checkpointIndex);
          } else if (functionCall.name === "convert_text_to_word_math") {
            const result = await executeConvertTextToWordMath(
              args.targetText || "",
              args.latex || "",
              args.scope || "selection",
              !!args.replaceAll,
              args.targets || null
            );
            toolResult = result.message;
            toolSucceeded = !!result.showToUser;

            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: "convert_text_to_word_math",
              result: toolResult,
              success: result.showToUser
            });

            updateSystemMessage(loadingMsg, toolResult, result.checkpointIndex);
          } else if (functionCall.name === "run_word_script") {
            const result = await executeRunWordScript(
              args.operations || [],
              args.description || "",
              args.javascript || ""
            );
            toolResult = result.message;
            toolSucceeded = !!result.success;

            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: args.description || "run_word_script",
              result: toolResult,
              success: result.success
            });

            updateSystemMessage(loadingMsg, toolResult, result.checkpointIndex);
          } else if (functionCall.name === "navigate_to_section") {
            updateSystemMessage(loadingMsg, `Navigating to: "${instruction}"...`);
            toolResult = await executeNavigate(instruction, docText);
            toolSucceeded = true;

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: true
            });

            updateSystemMessage(loadingMsg, `Navigated to: "${instruction}"`);
          } else if (functionCall.name === "edit_list") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Editing list from P${args.startParagraphIndex} to P${args.endParagraphIndex}...`);

            const result = await executeEditList(
              args.startParagraphIndex,
              args.endParagraphIndex,
              args.newItems,
              args.listType,
              args.numberingStyle
            );
            toolResult = result.message;
            toolSucceeded = !!result.success;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `edit_list P${args.startParagraphIndex}-P${args.endParagraphIndex}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          } else if (functionCall.name === "insert_list_item") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Inserting list item after P${args.afterParagraphIndex}...`);

            const result = await executeInsertListItem(
              args.afterParagraphIndex,
              args.text,
              args.indentLevel || 0
            );
            toolResult = result.message;
            toolSucceeded = !!result.success;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `insert_list_item after P${args.afterParagraphIndex}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          } else if (functionCall.name === "edit_table") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Editing table (${args.action})...`);

            const result = await executeEditTable(
              args.paragraphIndex,
              args.action,
              args.content,
              args.targetRow,
              args.targetColumn
            );
            toolResult = result.message;
            toolSucceeded = !!result.success;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `edit_table at P${args.paragraphIndex}: ${args.action}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          } else if (functionCall.name === "edit_section") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Editing section at P${args.sectionHeaderIndex}...`);

            const result = await executeEditSection(
              args.sectionHeaderIndex,
              args.newHeaderText,
              args.newBodyParagraphs,
              args.preserveSubsections
            );
            toolResult = result.message;
            toolSucceeded = !!result.success;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `edit_section at P${args.sectionHeaderIndex}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          } else if (functionCall.name === "convert_headers_to_list") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Converting ${args.paragraphIndices?.length || 0} headers to numbered list...`);

            const result = await executeConvertHeadersToList(
              args.paragraphIndices,
              args.newHeaderTexts,
              args.numberingFormat
            );
            toolResult = result.message;
            toolSucceeded = !!result.success;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `convert_headers_to_list: ${args.paragraphIndices?.join(', ')}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          }

          const isMutatingTool = mutatingToolNames.has(functionCall.name);
          if (isMutatingTool) {
            attemptedMutatingToolsThisLoop++;
            if (toolSucceeded) {
              successfulMutatingToolsThisLoop++;
            } else {
              let argsSignature = "";
              try {
                argsSignature = JSON.stringify(args || {});
              } catch (_) {
                argsSignature = "[unserializable-args]";
              }
              failedMutationSignatures.push(`${functionCall.name}|${argsSignature}|${toolResult || ""}`);
            }
          }

          // Move loading message to bottom after tool output
          if (loadingMsg) {
            const chatMessages = document.getElementById("chat-messages");
            if (chatMessages) chatMessages.appendChild(loadingMsg);
          }

          // Collect this function response

          // Shape this exactly as Gemini expects:
          // functionResponse: {
          //   name: "tool_name",
          //   response: {
          //     name: "tool_name",
          //     content: [ { text: "..." } ]
          //   }
          // }
          const functionResponse = {
            functionResponse: {
              name: functionCall.name,
              response: {
                name: functionCall.name,
                content: [
                  {
                    text: toolResult || ""
                  }
                ]
              },
              id: functionCall.id
            }
          };
          if (!functionCall.id) {
            delete functionResponse.functionResponse.id;
          }
          functionResponses.push(functionResponse);
        }

        // NOW add both the model's function call and the responses to history together
        // This ensures they're added as a complete pair
        chatHistory.push({
          role: "model",
          parts: parts
        });

        chatHistory.push({
          role: "user",
          parts: functionResponses
        });

        if (attemptedMutatingToolsThisLoop > 0 && successfulMutatingToolsThisLoop === 0) {
          const noProgressSignature = failedMutationSignatures.join("||").slice(0, 2000);
          const signatureChanged = !!(lastNoProgressSignature && noProgressSignature && noProgressSignature !== lastNoProgressSignature);
          consecutiveNoProgressToolLoops++;
          lastNoProgressSignature = noProgressSignature;

          console.warn(
            `[LoopGuard] No-progress mutation loop ${consecutiveNoProgressToolLoops}/${DOCUMENT_LIMITS.MAX_NO_PROGRESS_TOOL_LOOPS}`
              + (signatureChanged ? " (signature changed)" : "")
          );

          if (consecutiveNoProgressToolLoops >= DOCUMENT_LIMITS.MAX_NO_PROGRESS_TOOL_LOOPS) {
            const failureSummary = failedMutationSignatures
              .map(signature => signature.split("|").slice(-1)[0])
              .filter(Boolean)
              .slice(0, 2)
              .join(" ");
            const loopGuardMessage = failureSummary
              ? `I could not apply that change to the Word document. ${failureSummary}`
              : "I could not apply that change to the Word document with the available tools.";
            if (loadingMsg) {
              removeMessage(loadingMsg);
            } else {
              removeMessage(loadingMsg);
            }
            addMessageToChat("Gemini", loopGuardMessage);
            chatHistory.push({
              role: "model",
              parts: [{ text: loopGuardMessage }]
            });
            // Reset conversation history to avoid carrying forward orphaned
            // function-call/function-response turns into the next request.
            chatHistory = [];
            keepLooping = false;
            break;
          }
        } else {
          consecutiveNoProgressToolLoops = 0;
          lastNoProgressSignature = "";
        }

      } else {
        // Normal text response - this ends the loop
        // Robustly find the text part, skipping thought/thinking parts for the UI
        const textPart = parts.find(p => p.text && !p.thought);
        const aiResponse = textPart ? textPart.text : "Response generated (see document for changes).";

        // Add model response to history with proper structure
        chatHistory.push({
          role: "model",
          parts: parts
        });

        if (toolsExecutedInCurrentRequest.length === 0) {
          removeMessage(loadingMsg);
        }
        addMessageToChat("Gemini", aiResponse);
        keepLooping = false;
      }
    }

    // Maintain rolling window - but ensure we don't break function call/response pairs
    if (chatHistory.length > 10) {
      chatHistory = maintainHistoryWindow(chatHistory, 10);
    }

  } catch (error) {
    console.error("Error calling Gemini API:", error);

    // Handle user cancellation specifically
    if (error.message === 'Request cancelled by user') {
      removeMessage(loadingMsg);
      addMessageToChat("System", "Request cancelled.");
    } else {
      // Only remove loadingMsg if no tools were executed (meaning it's still a "Thinking" message)
      if (toolsExecutedInCurrentRequest.length === 0) {
        removeMessage(loadingMsg);

        // Cleanup history for failed requests (timeout or error)
        if (chatHistory.length > 0 && chatHistory[chatHistory.length - 1].role === "user") {
          console.log("Discarding failed request from history");
          chatHistory.pop();
        }
      }

      let errorMessage = error.message ? `Sorry, I couldn't get a response. Error: ${error.message}` : `Sorry, I couldn't get a response. Error: ${String(error)}`;

      // Override error message for timeouts
      if (error.message && (error.message.includes("timed out") || error.message.includes("timeout"))) {
        errorMessage = "The request failed before any document changes were applied. Gemini 2.5 Flash is the recommended fast model for document edits.";
      }

      const errorMsgEl = addMessageToChat("Error", errorMessage);

      // Add retry button if it's the specific missing content error
      if (error.message && error.message.includes("Gemini response was missing content.parts")) {
        addRetryButton(errorMsgEl, userMessage);
      }
    }
  } finally {
    // Clear the global abort controller
    currentRequestController = null;

    // Unlock UI
    chatInput.disabled = false;
    sendButton.disabled = false;
    if (thinkButton) thinkButton.disabled = false;
    chatInput.focus();
  }
}

// Helper with retry logic and timeout support
async function callGeminiWithModelFallback(apiKey, preferredModel, payload) {
  const normalizedPreferred = normalizeModelName(preferredModel);
  const preferredUrl = `https://generativelanguage.googleapis.com/v1beta/models/${normalizedPreferred}:generateContent?key=${apiKey}`;

  try {
    return await callGeminiWithRetry(preferredUrl, payload);
  } catch (error) {
    if (!normalizedPreferred || normalizedPreferred === EDIT_FALLBACK_MODEL) {
      throw error;
    }
    console.warn(`Primary model ${normalizedPreferred} failed; retrying with ${EDIT_FALLBACK_MODEL}.`, error);
    const fallbackUrl = `https://generativelanguage.googleapis.com/v1beta/models/${EDIT_FALLBACK_MODEL}:generateContent?key=${apiKey}`;
    return callGeminiWithRetry(fallbackUrl, payload);
  }
}

async function callGeminiWithRetry(url, payload, retries = 3, backoff = 1000) {
  for (let i = 0; i < retries; i++) {
    // Create abort controller for this specific fetch attempt
    const fetchController = new AbortController();

    // Create timeout that will abort the fetch
    const timeoutId = setTimeout(() => {
      fetchController.abort();
    }, TIMEOUT_LIMITS.FETCH_TIMEOUT_MS);

    try {
      // Also check if the global request controller was aborted (user cancelled)
      if (currentRequestController && currentRequestController.signal.aborted) {
        throw new Error('Request cancelled by user');
      }

      // Listen for global cancellation
      const onGlobalAbort = () => fetchController.abort();
      if (currentRequestController) {
        currentRequestController.signal.addEventListener('abort', onGlobalAbort);
      }

      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
        signal: fetchController.signal
      });

      // Clean up listeners
      clearTimeout(timeoutId);
      if (currentRequestController) {
        currentRequestController.signal.removeEventListener('abort', onGlobalAbort);
      }

      if (!response.ok) {
        const text = await response.text();

        // Check for the specific function call/response error (400 error)
        const isFunctionCallError = response.status === 400 &&
          text.includes("function response turn comes immediately after a function call turn");

        if (isFunctionCallError) {
          // Don't retry this error here - let the caller handle it
          throw new Error(`API failed: ${text}`);
        }

        // Only retry on 5xx errors
        if (response.status >= 500 && response.status < 600) {
          console.warn(`Attempt ${i + 1} failed with ${response.status}: ${text}`);
          if (i === retries - 1) throw new Error(`API failed after ${retries} attempts: ${text}`);
          // Wait before retrying
          await new Promise(r => setTimeout(r, backoff * Math.pow(2, i))); // Exponential backoff
          continue;
        }

        throw new Error(`API failed: ${text}`);
      }

      return await response.json();
    } catch (error) {
      clearTimeout(timeoutId);

      // Check if this was a user cancellation
      if (error.name === 'AbortError' || error.message === 'Request cancelled by user') {
        if (currentRequestController && currentRequestController.signal.aborted) {
          throw new Error('Request cancelled by user');
        }
        // This was a timeout abort
        console.warn(`Attempt ${i + 1} timed out after ${TIMEOUT_LIMITS.FETCH_TIMEOUT_MS / 1000}s`);
        if (i === retries - 1) {
          throw new Error(`Request timed out. The AI is taking longer than usual. Please try again.`);
        }
        await new Promise(r => setTimeout(r, backoff * Math.pow(2, i)));
        continue;
      }

      // If it's the function call error, throw immediately without retry
      if (error.message && error.message.includes("function response turn comes immediately after a function call turn")) {
        throw error;
      }

      if (i === retries - 1) throw error;
      console.warn(`Attempt ${i + 1} failed: ${error.message}`);
      await new Promise(r => setTimeout(r, backoff * Math.pow(2, i)));
    }
  }
}

function flattenContentsForLive(contents) {
  return (contents || [])
    .map(turn => {
      const role = turn.role || "user";
      const parts = (turn.parts || []).map(part => {
        if (part.text) return part.text;
        if (part.functionCall) {
          return `[tool requested: ${part.functionCall.name} ${JSON.stringify(part.functionCall.args || {})}]`;
        }
        if (part.functionResponse) {
          return `[tool result: ${part.functionResponse.name} ${JSON.stringify(part.functionResponse.response || {})}]`;
        }
        return "";
      }).filter(Boolean).join("\n");
      return parts ? `${role.toUpperCase()}:\n${parts}` : "";
    })
    .filter(Boolean)
    .join("\n\n");
}

function livePartToGeneratePart(part) {
  if (part.text) return { text: part.text };
  if (part.functionCall) return { functionCall: part.functionCall };
  return null;
}

async function callGeminiLiveAsGenerateContent(apiKey, modelName, payload) {
  return new Promise((resolve, reject) => {
    const wsUrl = `wss://generativelanguage.googleapis.com/ws/google.ai.generativelanguage.v1beta.GenerativeService.BidiGenerateContent?key=${encodeURIComponent(apiKey)}`;
    const websocket = new WebSocket(wsUrl);
    const parts = [];
    let settled = false;
    let setupDone = false;
    let outputText = "";
    let clientContentSent = false;

    const sendClientContent = () => {
      if (clientContentSent || websocket.readyState !== WebSocket.OPEN) return;
      clientContentSent = true;
      websocket.send(JSON.stringify({
        clientContent: {
          turns: [{
            role: "user",
            parts: [{ text: flattenContentsForLive(payload.contents) }]
          }],
          turnComplete: true
        }
      }));
    };

    const timeoutId = setTimeout(() => {
      if (!settled) {
        settled = true;
        try { websocket.close(); } catch (_) { /* ignore */ }
        reject(new Error("Live API request timed out."));
      }
    }, TIMEOUT_LIMITS.FETCH_TIMEOUT_MS);

    const finish = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timeoutId);
      try { websocket.close(); } catch (_) { /* ignore */ }
      if (parts.length === 0 && outputText.trim()) {
        parts.push({ text: outputText.trim() });
      }
      if (parts.length === 0) {
        parts.push({ text: "Live API response completed. See document for any applied tool changes." });
      }
      resolve({
        candidates: [{
          content: {
            role: "model",
            parts
          }
        }]
      });
    };

    websocket.onerror = () => {
      if (settled) return;
      settled = true;
      clearTimeout(timeoutId);
      reject(new Error("Live API WebSocket failed."));
    };

    websocket.onopen = () => {
      const setupMessage = {
        setup: {
          model: `models/${normalizeModelName(modelName)}`,
          generationConfig: {
            responseModalities: ["AUDIO"]
          },
          systemInstruction: payload.systemInstruction,
          tools: payload.tools
        }
      };
      websocket.send(JSON.stringify(setupMessage));
      setTimeout(sendClientContent, 1000);
    };

    websocket.onmessage = async (event) => {
      const rawData = event.data && typeof event.data.text === "function"
        ? await event.data.text()
        : event.data;
      const response = JSON.parse(rawData);

      if (!setupDone && response.setupComplete !== undefined) {
        setupDone = true;
        sendClientContent();
        return;
      }

      if (response.toolCall?.functionCalls?.length) {
        for (const functionCall of response.toolCall.functionCalls) {
          parts.push({
            functionCall: {
              name: functionCall.name,
              args: functionCall.args || {},
              id: functionCall.id
            }
          });
        }
        finish();
        return;
      }

      if (response.serverContent?.modelTurn?.parts?.length) {
        for (const part of response.serverContent.modelTurn.parts) {
          const mapped = livePartToGeneratePart(part);
          if (mapped) parts.push(mapped);
        }
      }

      if (response.serverContent?.outputTranscription?.text) {
        outputText += response.serverContent.outputTranscription.text;
      }

      if (response.serverContent?.turnComplete || response.serverContent?.generationComplete) {
        finish();
      }
    };

    if (currentRequestController) {
      currentRequestController.signal.addEventListener("abort", () => {
        if (!settled) {
          settled = true;
          clearTimeout(timeoutId);
          try { websocket.close(); } catch (_) { /* ignore */ }
          reject(new Error("Request cancelled by user"));
        }
      }, { once: true });
    }
  });
}

