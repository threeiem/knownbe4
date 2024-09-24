function processInbox() {
  var pageToken;
  var processedCount = 0;
  const maxToProcess = 500; // Adjust as needed
  
  do {
    try {
      var threads = GmailApp.search("newer_than:7d", 0, 100);
      
      for (var i = 0; i < threads.length && processedCount < maxToProcess; i++) {
        var messages = threads[i].getMessages();
        for (var j = 0; j < messages.length && processedCount < maxToProcess; j++) {
          processMessage(messages[j]);
          processedCount++;
        }
      }
      
      if (threads.length < 100) {
        break;
      }
      
      pageToken = threads[threads.length - 1].getId();
    } catch (e) {
      console.error('Error processing inbox:', e);
      break;
    }
  } while (processedCount < maxToProcess);
  
  console.log('Processed ' + processedCount + ' messages');
}

function processMessage(message) {
  try {
    var header = message.getHeader("X-PHISHTEST");
    if (header == "This is a phishing security test from KnowBe4 that has been authorized by the recipient organization") {
      var label = GmailApp.getUserLabelByName("knowbe4");
      if (!label) {
        label = GmailApp.createLabel("knowbe4");
        console.log('Created new label: knowbe4');
      }
      var thread = message.getThread();
      thread.addLabel(label);
      archiveThread(thread);
      console.log('Labeled and archived message: ' + message.getSubject());
    }
  } catch (e) {
    console.error('Error processing message:', e);
  }
}

function archiveThread(thread) {
  thread.moveToArchive();
  
  // Double-check if the thread is still in the inbox
  Utilities.sleep(1000); // Wait for 1 second
  if (thread.isInInbox()) {
    console.log('Thread still in inbox after first attempt, trying again...');
    thread.moveToArchive();
    
    // Final check
    Utilities.sleep(1000);
    if (thread.isInInbox()) {
      console.error('Failed to archive thread: ' + thread.getFirstMessageSubject());
    } else {
      console.log('Successfully archived on second attempt: ' + thread.getFirstMessageSubject());
    }
  } else {
    console.log('Successfully archived: ' + thread.getFirstMessageSubject());
  }
}

function main() {
  console.log('Starting inbox processing...');
  processInbox();
  console.log('Finished inbox processing.');
}