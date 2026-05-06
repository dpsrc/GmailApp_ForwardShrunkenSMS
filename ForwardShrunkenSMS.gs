// Destination and Alias configuration
const targetEmail = "put yours here";
const aliasEmail = "put yours here"; // Must be added in Gmail Settings -> Accounts and Import
const labelInputName = "put yours here";
const labelOutputName = "put yours here";

/**
 * Fully automated email forwarder with text replacement.
 * Features: Alias support, state management, and Plain body processing.
 */
function processAndForwardEmails() {
  const userProperties = PropertiesService.getUserProperties();
  const lastCheck = userProperties.getProperty('LAST_CHECK_TIME');
  console.log("lastCheck = " + lastCheck + " = " + Utilities.formatDate(new Date(parseInt(lastCheck)), "UTC", "yyyy-MM-dd HH:mm:ss.sss 'UTC'"));
  const now = new Date();
  
  // Search query: looking for emails in 'ToForward' label
  // Gmail search 'after' works with seconds since epoch
  let query = "label:" + labelInputName + " from:('voice-noreply@google.com' OR 'txt.voice.google.com')";
  if (lastCheck) {
    console.log("Entered if (lastCheck)");
    const secondsSinceEpoch = Math.floor(lastCheck / 1000);
    query += " after:" + secondsSinceEpoch;
    console.log("query = " + query);
  
    const threads = GmailApp.search(query);

    threads.forEach(thread => {
      console.log("Entered forEach(thread =>");
      const messages = thread.getMessages();
      messages.forEach(message => {
        const messageDate = message.getDate().getTime();

        console.log("Subject = " + message.getSubject());
        
        // Ensure we only process messages newer than our last stored timestamp
        if (!lastCheck || messageDate > lastCheck) {
          console.log("Entered if (!lastCheck || messageDate > lastCheck)");
          let smsText = extractVoiceText(message.getPlainBody());

          // Send the modified email
          sendAndDeleteEmail(targetEmail, processSubject(message.getSubject()).replace("New text message from", "SMS from"),
              smsText, aliasEmail, labelOutputName);
        }
      });
    })
  };

  // Save the current timestamp for the next execution cycle
  userProperties.setProperty('LAST_CHECK_TIME', now.getTime().toString());
}

function extractVoiceText(plainBodyText) {
// Regex explanation:
  // <https:\/\/voice\.google\.com>  - Matches the starting URL.
  // \s*                             - Matches any leading whitespace/newlines.
  // ([\s\S]*?)                      - Capturing Group: matches ANY character (including newlines) 
  //                                   non-greedily until the next part of the regex is found.
  // \s*                             - Matches any trailing whitespace.
  // (?: ... | ... )                 - Non-capturing group with two alternatives:
  //    Option A: "YOUR ACCOUNT <https://voice.google.com>"
  //    Option B: "To respond to this message, launch Google Voice"
  const regex = /<https:\/\/voice\.google\.com>\s*([\s\S]*?)\s*(?:YOUR ACCOUNT <https:\/\/voice\.google\.com>|To respond to this message, launch Google Voice)/;
  
  const match = plainBodyText.match(regex);
  
  if (match && match[1]) {
    const extracted = match[1];
    console.log("Extracted text: " + extracted);
    return extracted;
  } else {
    console.log("No match found.");
    return plainBodyText;
  }
}

/**
 * Processes the email subject where the name can consist of multiple words.
 * @param {string} subject - The original email subject.
 * @return {string} - The modified subject.
 */
function processSubject(subject) {
  // Regex explanation:
  // (New text message from [\s\wÀ-ßà-ÿ¨¸²³¯¿ªº¥´]+?) - Group 1: Matches prefix and name.
  //                                                   [\s\w...] allows spaces, letters, and Cyrillic.
  //                                                   The '?' makes it non-greedy so it stops at the first number.
  // \s+                                              - Matches the space before the number.
  // ([\d\(\s\-\)]+)                                  - Group 2: Matches the phone number pattern.
  const multiWordPattern = /(New text message from [\s\w\u0400-\u04FF]+?)\s+([\d\(\s\-\)]+)/;

  const match = subject.match(multiWordPattern);

  if (match) {
    // Returns the part with the name, trimming any trailing spaces
    return match[1].trim();
  }

  // Returns original if no name+number pattern is found
  return subject;
}

/**
 * Sends an email and applies a label to the sent thread.
 */
function sendAndDeleteEmail(targetEmail, subject, body, aliasEmail, labelName) {
  // Send the email
  // Note: sendEmail does not return a message object in Apps Script
  GmailApp.sendEmail(targetEmail, subject, body, { 
    from: aliasEmail
  });

  // A small delay to ensure Gmail indexes the message in the Sent folder
  Utilities.sleep(1500);

  // Find the thread we just sent
  // We search in Sent messages with the specific subject
  const threads = GmailApp.search('in:sent subject:"' + subject + '"', 0, 1);

  if (threads.length > 0) {
    const thread = threads[0];
    
    // Get or create the label
    let label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      label = GmailApp.createLabel(labelName);
      console.log("Created new label: " + labelName);
    }

    // Apply the label to the thread
    thread.addLabel(label);
    console.log("Label '" + labelName + "' applied to the sent thread.");

    // Move the thread to Trash after labeling
    thread.moveToTrash();
    console.log("Thread moved to Trash.");
  } else {
    console.log("Could not find the sent email to apply the label and move to trash.");
  }
}
