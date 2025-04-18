/**
 * @OnlyCurrentDoc
 *
 * This script provides a custom function (=chat1) for Google Sheets
 * to interact with various AI models (Gemini, ChatGPT, DeepSeek)
 * and includes a configuration sidebar to store API keys securely.
 *
 * To use:
 * 1. Open your Google Sheet or Doc.
 * 2. Go to Extensions > Apps Script.
 * 3. Replace any existing code with this entire script.
 * 4. Save the script.
 * 5. Refresh your Sheet/Doc.
 * 6. Go to the "AI Functions" menu > "Configure API Keys" to enter your keys.
 * 7. Use the =chat1() function in your Sheet.
 */

// --- API Configuration Storage ---

/**
 * Gets the user properties where API keys are stored.
 * @return {Properties} The user properties object.
 */
function getUserProperties() {
  return PropertiesService.getUserProperties();
}

/**
 * Saves the API keys to user properties.
 * This function is called from the sidebar client-side script.
 * @param {Object} apiKeys An object containing the API keys (gemini, chatgpt, deepseek).
 */
function saveApiKeys(apiKeys) {
  const userProperties = getUserProperties();
  userProperties.setProperty('geminiApiKey', apiKeys.gemini || '');
  userProperties.setProperty('chatgptApiKey', apiKeys.chatgpt || '');
  userProperties.setProperty('deepseekApiKey', apiKeys.deepseek || '');
  console.log('API keys saved.');
}

/**
 * Retrieves the API keys from user properties.
 * This function is called from the sidebar client-side script.
 * @return {Object} An object containing the stored API keys.
 */
function loadApiKeys() {
  const userProperties = getUserProperties();
  return {
    gemini: userProperties.getProperty('geminiApiKey'),
    chatgpt: userProperties.getProperty('chatgptApiKey'),
    deepseek: userProperties.getProperty('deepseekApiKey')
  };
}

// --- API Endpoints ---

const API_URLS = {
  gemini: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=',
  // Placeholder URL for ChatGPT - replace if needed based on your setup
  chatgpt: 'https://api.openai.com/v1/chat/completions',
  deepseek: 'https://api.deepseek.com/v1/chat/completions'
};

// --- Core API Call Function ---

/**
 * Calls the specified AI API to generate content based on a prompt and optional context.
 * @param {string} apiType The type of API to use ('gemini', 'chatgpt', or 'deepseek').
 * @param {string} prompt The main prompt for the language model.
 * @param {string} [context] Optional context string to include in the prompt.
 * @return {string} The generated text response, or an error message.
 */
function callApi(apiType, prompt, context) {
  const lowerApiType = apiType.toLowerCase();
  const apiKeys = loadApiKeys();
  let apiKey;
  let apiUrl;
  let options;
  let fullPrompt = prompt;

  // Prepend context if provided
  if (context && typeof context === 'string' && context.length > 0) {
      fullPrompt = `Context: ${context}\n\nPrompt: ${prompt}`;
  }

  // Determine API key and URL based on type
  switch (lowerApiType) {
    case 'gemini':
      apiKey = apiKeys.gemini;
      apiUrl = API_URLS.gemini + apiKey;
      if (!apiKey) {
        throw new Error('Gemini API Key is not configured. Please set it via the "AI Functions" menu.');
      }
      options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify({
          contents: [{
            parts: [{
              text: fullPrompt
            }]
          }]
        })
      };
      break;
    case 'chatgpt':
      apiKey = apiKeys.chatgpt;
      apiUrl = API_URLS.chatgpt;
       if (!apiKey) {
        throw new Error('ChatGPT API Key is not configured. Please set it via the "AI Functions" menu.');
      }
      // ChatGPT API expects a different payload structure
      options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
          'Authorization': 'Bearer ' + apiKey
        },
        'payload': JSON.stringify({
          model: "gpt-3.5-turbo", // You can change the model if needed
          messages: [{
            role: "user",
            content: fullPrompt
          }],
          stream: false // Set to true for streaming, but requires different handling
        })
      };
      break;
    case 'deepseek':
      apiKey = apiKeys.deepseek;
      apiUrl = API_URLS.deepseek;
       if (!apiKey) {
        throw new Error('DeepSeek API Key is not configured. Please set it via the "AI Functions" menu.');
      }
      // DeepSeek API expects a different payload structure
      options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
          'Authorization': 'Bearer ' + apiKey
        },
        'payload': JSON.stringify({
          model: "deepseek-coder", // You can change the model if needed
          messages: [{
            role: "user",
            content: fullPrompt
          }],
          stream: false // Set to true for streaming, but requires different handling
        })
      };
      break;
    default:
      return 'Error: Invalid API type specified. Use "g", "c", or "d".';
  }


  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const data = JSON.parse(response.getContentText());

    // Parse response based on API type
    switch (lowerApiType) {
      case 'gemini':
        if (data && data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0]) {
          return data.candidates[0].content.parts[0].text;
        } else {
          console.error('Unexpected Gemini API response structure:', JSON.stringify(data));
          return 'Error: Unable to get a valid response from the Gemini API. Check the script logs.';
        }
      case 'chatgpt':
         if (data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) {
          return data.choices[0].message.content;
        } else {
           console.error('Unexpected ChatGPT API response structure:', JSON.stringify(data));
          return 'Error: Unable to get a valid response from the ChatGPT API. Check the script logs.';
        }
      case 'deepseek':
         if (data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) {
          return data.choices[0].message.content;
        } else {
           console.error('Unexpected DeepSeek API response structure:', JSON.stringify(data));
          return 'Error: Unable to get a valid response from the DeepSeek API. Check the script logs.';
        }
      default:
        // This case should ideally not be reached due to the switch above
        return 'Error: Unknown API type encountered during response parsing.';
    }

  } catch (e) {
    console.error(`API call failed for ${apiType}:`, e);
    return `Error: ${apiType} API call failed. See script logs for details.`;
  }
}

// --- Custom Google Sheets Function ---

/**
 * Custom function for Google Sheets to generate text using a specified AI API.
 * Use in a cell like:
 * =chat1("g", "Write a short poem about nature.") // Uses Gemini
 * =chat1("c", "Write a Python function to calculate factorial.") // Uses ChatGPT
 * =chat1("d", "Explain quantum computing in simple terms.") // Uses DeepSeek
 * =chat1("g", "Summarize this.", A1) // Uses Gemini, includes content of A1 as context
 *
 * @param {string} modelCode The code for the AI model ('g', 'c', or 'd').
 * @param {string} prompt The prompt for the language model.
 * @param {string} [contextCellReference] Optional cell reference (e.g., "A1") for context.
 * @return {string} The generated text.
 * @customfunction
 */
function chat1(modelCode, prompt, contextCellReference) {
  let apiType;
  switch (modelCode.toLowerCase()) {
    case 'g':
      apiType = 'gemini';
      break;
    case 'c':
      apiType = 'chatgpt';
      break;
    case 'd':
      apiType = 'deepseek';
      break;
    default:
      return 'Error: Invalid model code. Use "g", "c", or "d".';
  }

  if (typeof prompt !== 'string' || prompt.length === 0) {
    return "Error: Prompt must be a non-empty string.";
  }

  let context = '';
  if (contextCellReference) {
    try {
      // Get the value from the referenced cell
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      // Ensure contextCellReference is treated as a string cell reference
      const range = sheet.getRange(String(contextCellReference));
      context = range.getValue();
      // Convert context to string in case it's a number, date, etc.
      if (context !== null && context !== undefined) {
          context = String(context);
      } else {
          context = ''; // Handle empty or null cell values
      }
    } catch (e) {
      console.error('Error reading context cell:', e);
      return `Error reading context cell ${contextCellReference}: ${e.message}`;
    }
  }

  return callApi(apiType, prompt, context);
}

// --- UI Functions ---

/**
 * Adds a custom menu to the Google Sheet or Doc UI.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi() || DocumentApp.getUi();
  ui.createMenu('AI Functions')
    .addItem('Configure API Keys', 'showConfigSidebar')
    .addItem('How to use chat1', 'showChat1Info')
    .addToUi();
}

/**
 * HTML content for the configuration sidebar.
 * Embedded as a string for a single-file script.
 */
const configSidebarHtml = `
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      /* Basic styling for the sidebar */
      body {
        font-family: sans-serif;
        margin: 10px;
      }
      label {
        display: block;
        margin-bottom: 5px;
        font-weight: bold;
      }
      input[type="text"] {
        width: 95%;
        padding: 5px;
        margin-bottom: 10px;
        border: 1px solid #ccc;
        border-radius: 4px;
      }
      button {
        background-color: #4CAF50; /* Green */
        border: none;
        color: white;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        font-size: 16px;
        margin-top: 10px;
        cursor: pointer;
        border-radius: 4个体;
      }
       button:hover {
        background-color: #45a049;
      }
      .success {
          color: green;
          margin-top: 10px;
      }
       .error {
          color: red;
          margin-top: 10px;
      }
    </style>
  </head>
  <body>
    <h2>Configure API Keys</h2>
    <p>Enter your API keys below. They will be stored securely with your script.</p>

    <label for="geminiApiKey">Gemini API Key:</label>
    <input type="text" id="geminiApiKey" name="geminiApiKey"><br>

    <label for="chatgptApiKey">ChatGPT API Key:</label>
    <input type="text" id="chatgptApiKey" name="chatgptApiKey"><br>

    <label for="deepseekApiKey">DeepSeek API Key:</label>
    <input type="text" id="deepseekApiKey" name="deepseekApiKey"><br>

    <button onclick="saveKeys()">Save Keys</button>

    <div id="status"></div>

    <script>
      // Load existing keys when the sidebar opens
      // Use google.script.run to call server-side functions
      google.script.run.withSuccessHandler(loadExistingKeys).loadApiKeys();

      function loadExistingKeys(apiKeys) {
        document.getElementById('geminiApiKey').value = apiKeys.gemini || '';
        document.getElementById('chatgptApiKey').value = apiKeys.chatgpt || '';
        document.getElementById('deepseekApiKey').value = apiKeys.deepseek || '';
      }

      function saveKeys() {
        const geminiKey = document.getElementById('geminiApiKey').value;
        const chatgptKey = document.getElementById('chatgptApiKey').value;
        const deepseekKey = document.getElementById('deepseekApiKey').value;

        const apiKeys = {
          gemini: geminiKey,
          chatgpt: chatgptKey,
          deepseek: deepseekKey
        };

        const statusDiv = document.getElementById('status');
        statusDiv.textContent = 'Saving...';
        statusDiv.className = ''; // Clear previous status classes

        // Use google.script.run to call server-side save function
        google.script.run
            .withSuccessHandler(function() {
                statusDiv.textContent = 'Keys saved successfully!';
                statusDiv.className = 'success';
            })
            .withFailureHandler(function(error) {
                 statusDiv.textContent = 'Error saving keys: ' + error.message;
                 statusDiv.className = 'error';
                 console.error('Error saving keys:', error);
            })
            .saveApiKeys(apiKeys);
      }
    </script>
  </body>
</html>
`; // End of configSidebarHtml string

/**
 * Opens the API key configuration sidebar using the embedded HTML.
 */
function showConfigSidebar() {
  const htmlOutput = HtmlService.createHtmlOutput(configSidebarHtml)
    .setTitle('Configure API Keys')
    .setWidth(300); // Adjust width as needed
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}


/**
 * Shows information about using the chat1 custom function in Sheets.
 */
function showChat1Info() {
  const ui = SpreadsheetApp.getUi() || DocumentApp.getUi();
  ui.alert(
    'How to use the chat1 function',
    'Use the custom function =chat1("model_code", "Your prompt here", "OptionalCellReference").\n\n' +
    'model_code: "g" for Gemini, "c" for ChatGPT, "d" for DeepSeek.\n' +
    'OptionalCellReference: A cell like A1, B2, etc., whose content will be included as context.\n\n' +
    'Examples:\n' +
    '=chat1("g", "Write a short poem about nature.")\n' +
    '=chat1("c", "Write a Python function to calculate factorial.")\n' +
    '=chat1("d", "Explain quantum computing in simple terms.")\n' +
    '=chat1("g", "Summarize this content.", "A1")\n\n' +
    'Configure your API keys using the "AI Functions" > "Configure API Keys" menu.',
    ui.ButtonSet.OK
  );
}

// Note: This script primarily focuses on the Sheets custom function.
// To add Docs functionality (e.g., processing selected text), you would
// add new menu items and functions that read selected text using
// DocumentApp.getActiveDocument().getSelection().
