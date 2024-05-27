/******************************************************************************************************
 *
 * Name:                 generateActionItems
 * Description:          This Google Apps script is intended to be used in Doogle Doc to call the
 *                       OpenAI API: https://api.openai.com/v1/chat/completions to generate a formatted
 *                       list of action items extracted from meeting minutes, and insert them into
 *                       the Google Doc.
 * Date:                 May 25, 2024
 * Author:               Mark Stankevicius
 * GitHub Repository:    https://github.com/stankev/chatgpt-googledoc-action-items-script
 *
 *******************************************************************************************************
 */
function generateActionItems() {

  // Define the constants for the API call to chat/completions and the model configuration
  // Replace the placeholder value of the const below with your API key 
  //const API_KEY = 'sk-nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn';     // replace this string with your OpenAI API key
 
  const API_URL = 'https://api.openai.com/v1/chat/completions';
  const MODEL_NAME = 'gpt-4-turbo-preview';                    // use a model that is compatible with chat/completions API - gpt-4o, gpt-4-turbo-preview, gpt-4, gpt3.5-turbo
  const TEMPERATURE = 0.2;
  const TOP_P = 1;
  const FREQUENCY_PENALTY = 0.0;    // use a value of 0.0 to direct the model to provide more accuracy and clarity in the generated output to remain true to the source 
  const PRESENCE_PENALTY = 0.0;     // use a value of 0.0 to ensure more accuracy by ensuring that the model maintains fidelity to the source text
 
  const SYSTEM_PROMPT = `Simulate three brilliant, logical project managers reviewing status meeting minutes and determining the action items.
    The three project managers review the provided status meeting minutes and create their list of action items. The three project managers must 
    carefully review the full list of meeting notes to ensure they capture any action items hidden in the minutes. 
    The three experts then compare their list against the action item list from the other project managers. 
    Based on their comparison they generate a final list of action items. Do not generate action items for milestones in the minutes.
    The action items should list the following for each item: title of the action item, due date if known, dependency if known, and owner of the action item. 
    If there is no dependency, then state None.  Try to infer the due date based on the minutes, but if the due date cannot be determined than specify TBD. 
    Be accurate when creating the action items and do not make up fictitious action items that are not in the meeting minutes.
    Format the output so that it could be inserted into a google doc. Do not bold text or any special formatting such as asterisks for emphasis.
    An example format of output is the following:
    Action Item: Develop Power User training plan and materials.\n - Due Date: April 25/2024\n  - Dependency: Completion of Teams training materials.\n - Owner: Alice Williams (Training Lead)`;

  // Extract the meeting minutes from the document
  const chatPromptMinutes = extractText('Meeting minutes start:', 'Meeting minutes end');    // make sure your doc has these start and end delimiters, including case sensitive matches
  // Validate that the meeting minutes were found. If meeting minutes are not found or empty send message to the log
  if (!chatPromptMinutes) {
    Logger.log('The delimiters were found, but the meeting minutes were empty');
    return null;
  }

  // create the payload for the chat/completions API request
  const payload = {
    model: MODEL_NAME,
    messages: [
      {
        role: 'system',
        content: SYSTEM_PROMPT
      },
      {
        role: 'user',
        content: chatPromptMinutes
      }
    ],
    max_tokens: 1000,
    temperature: TEMPERATURE,
    top_p: TOP_P,
    frequency_penalty: FREQUENCY_PENALTY,
    presence_penalty: PRESENCE_PENALTY
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),       // convert the payload object to a JSON string to send to chat/completions API
    headers: {
      Authorization: 'Bearer ' + API_KEY     // uses your API key that was assigned earlier in the code
    }
  };

  // call the chat/completions API to review the meeting minutes and generate action items
  try {
    const response = UrlFetchApp.fetch(API_URL, options);                      // Call to the Chat/completions API 
    const chatResponseJson = JSON.parse(response.getContentText());            // parse the response from OpenAI into a JSON object

    // Validate the response was successful and if successful append action items to document
    if (chatResponseJson && chatResponseJson.choices && chatResponseJson.choices.length > 0) {     // is there generated text in the response
    
      //Logger.log('chatResponseJson: ' + JSON.stringify(chatResponseJson, null, 2));    // log the JSON object to the apps script logger if needed for debugging later - uncomment this line if you need to debug
      
      const actionItems = chatResponseJson.choices[0].message.content;          // extract the action items from the content of the JSON generated by OpenAI chat/completions API
      DocumentApp.getActiveDocument().getBody().appendParagraph(actionItems);   // append the action items to the end of the document 
    } else {
      Logger.log('No action items were generated');
    }

  } catch (error) {                                                            // An error occurred when calling the chat/completions API so log the details
    Logger.log('An error occurred: '+ error.message);                          // log any errors to the apps script logger
    return null;
  }
}
  
  
/*
 * Create a custom menu in the Google Docs UI for the "AI Tools" menu with one item: "Generate Action Items".
 * When the "Generate Action Items" item is clicked, it calls the "generateActionITems" function.
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  const menu = ui.createMenu('AI Tools')
                 .addItem('Generate Action Items', 'generateActionItems');
  menu.addToUi();
}

/*  
 * Function to extract text between start and end delimiters to return the meeting minutes
 */
function extractText(startDelimiter, endDelimiter) {
  
  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();

  // Find the start and end strings positions
  const startPos = body.getText().indexOf(startDelimiter);
  const endPos = body.getText().indexOf(endDelimiter);

  // If both strings are found in the document
  if (startPos !== -1 && endPos !== -1 && startPos < endPos) {
    
    // Calculate the position after the start string and before the end string
    const startOfText = startPos + startDelimiter.length;
    const meetingMinutes = body.getText().substring(startOfText, endPos).trim();
    
    //Logger.log('first 25 characters of the meeting minutes: ' + meetingMinutes.substring(0, 25));   // log meeting minutes if needed during debugging problems - uncomment this line if you need to debug
    return meetingMinutes;
  
  } else {
    Logger.log('Start or end string not found, or start string is after end string.');
    return null;
  }

}
