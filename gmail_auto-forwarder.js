///////////////////////////////////////////////////////////////////////
// AVAILABLE FUNCTIONS FOR THE USER TO RUN                           //
// • 'preflightInspection': Validate the user's input manually.      //
// • 'setTrigger': Install a trigger for this script.                //
// • 'unsetTrigger': Uninstall all trigger kinds of this script.     //
// DO NOT RUN THESE FUNCTIONS MANUALLY. RUN THEM AT YOUR OWN RISK!!! //
// • 'waitExecution'                                                 //
// • 'waitPager'                                                     //
// • 'threadsPager'                                                  //
// • 'forwarder'                                                     //
// • 'forwarderPager'                                                //
// • 'gmailAutoForwarder'                                            //
///////////////////////////////////////////////////////////////////////

///////////////////
// CONFIGURATION //
///////////////////

// Sample list of filters and recipients
var PASS = [
  // The first element of the array contains the filter, and the last element is an array that stores the recipients
  ['from:"from_me@outlook.com" AND subject:my_outlook', ['to_you@yahoo.com', 'to_you@gmail.com']],
  ['from:"from_me@yahoo.com" AND subject:my_yahoo', ['to_you@outlook.com', 'to_you@gmail.com']],
  ['from:"from_me@gmail.com" AND subject:my_gmail', ['to_you@outlook.com', 'to_you@yahoo.com']]
]

// Specifies to run this script every [n] minute(s). [n] must be 1, 5, 10, 15 or 30
// DEFAULT: 5 minutes
var CLOCK = 5

// Specifies to delay batch in [n] minute(s)
// DEFAULT: 1 minute
var DELAY = 1

// Specifies to pause between the job on Inbox and Sent in [n] minute(s)
// DEFAULT: 0.002 minute
var PAUSE = 0.002

// Specifies how many messages to forward in one batch
// DEFAULT: 100 messages
var THREADS = 100

// Set 'true' to enable the complete log
// DEFAULT: false
var DEBUG = false

// Set 'true' to running 'preflightInspection' with Trial mode. This mode enabled the complete log and disabled forwarding and permanent deletion of matching message(s)
// DEFAULT: false
var TRIAL = false

////////////////////////////
// ADVANCED CONFIGURATION //
////////////////////////////

// Set 'true' in the first element to use an alias address and set '0' in the last element to use the first registered alias address
// DEFAULT: [false, 0]
var ALIAS = [false, 0]

// Set 'true' in the first element to wait for the running execution to end first, then in the last element, set this project Google Apps Script ID in the string format
// CAUTION: Configure the step 5 installation on README.md is required
// DEFAULT: [false, 'JjW34wcx7rTbwNCEPojdJ-THIS_IS_A_DUMMY_ID-BXSI69WdbS9rapVo']
var EXEC_WAIT = [false, 'JjW34wcx7rTbwNCEPojdJ-THIS_IS_A_DUMMY_ID-BXSI69WdbS9rapVo']

//////////////////////////////////////////////////////////////////
// AUTHOR  : ADE DESTRIANTO                                     //
// TITLE   : GMAIL AUTO-FORWARDER                               //
// VERSION : 1.3 BUILD 20240624                                 //
// GITHUB  : https://github.com/destrianto/gmail_auto-forwarder //
//////////////////////////////////////////////////////////////////

// Validate the user's input before setting a trigger for this script
function preflightInspection(mode = TRIAL, triggered = false){
  try{
    var line = []
    // Regex's reference: https://stackoverflow.com/a/201378
    var email = /(?:[a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])/
    // Iterate the list of filters and recipients to perform validation
    for(i = 0; i < PASS.length; i++){
      line.push([(i + 1), PASS[i]])
      GmailApp.search(PASS[i][0], 0, 1)
      for(j = 0; j < PASS[i][1].length; j++){
        if(!PASS[i][1][j].match(email)){
          throw 'email'
        }
      }
    }
    // Trial mode
    if(mode){
      console.info('  PRE-FLIGHT: The list of filters and recipients passed the validation inspection. Proceed to Trial mode')
      gmailAutoForwarder(false, mode)
      return
    }
    if(!triggered){
      console.info('PRE-FLIGHT: The list of filters and recipients passed the validation inspection. Run \'setTrigger\' to use the script')
    }
    // Continue this script execution
    return true
  }
  catch(caught){
    // Output appears if there's an invalid email address in the filter
    if(caught === 'email'){
      if(line.length < 2){
        console.error('FILTER_ERROR: Please revise the email address(es) in Filter No. ' + line[0][0])
      }
      else{
        console.error('FILTER_ERROR: Please revise the email address(es) in Filter No. ' + (line[line.length - 2][0] + 1) + '. The Filter No. ' + (line[line.length - 2][0] + 1) + ' is next to this filter:\n' + '[\'' + line[line.length - 2][1][0] + '\', [\'' + line[line.length - 2][1][1].join('\', \'') + '\']]')
      }
    }
    // Output appears if there's a missing comma in the filter
    else{
      var data = ['object', 'undefined']
      if(line.length < 2 && data.some(k => typeof line[0][1] === k)){
        console.error('FILTER_ERROR: Filter No. ' + line[0][0] + ' missed a comma separator at the end of its array or first element')
      }
      else{
        console.error('FILTER_ERROR: Filter No. ' + (line[line.length - 2][0] + 1) + ' missed a comma separator at the end of its array or first element. The Filter No. ' + (line[line.length - 2][0] + 1) + ' is next to this filter:\n' + '[\'' + line[line.length - 2][1][0] + '\', [\'' + line[line.length - 2][1][1].join('\', \'') + '\']]')
      }
    }
    // Stop this script execution at once if there's an invalid list
    return false
  }
}

// Set a clock trigger to this script
function setTrigger(){
  // Validate the user's input
  if(preflightInspection(false, true)){
    ScriptApp
      .newTrigger('forwarder')
      .timeBased()
      .everyMinutes(CLOCK)
      .create()
    console.info('The forwarder\'s clock trigger is adjusted every ' + CLOCK + ' minute(s)')
  }
}

// Unset trigger
function unsetTrigger(trigger = 'unset'){  
  var triggers = ScriptApp.getProjectTriggers()
  var handler = ['forwarder', 'forwarderPager', 'threadsPager']
  var found = 0
  for(i = 0; i < triggers.length; i++){
    // Unset specified trigger
    if(trigger !== 'unset'){
      if(triggers[i].getHandlerFunction() === trigger){
        ScriptApp.deleteTrigger(triggers[i])
      }
    }
    // Unset all kinds of this script's trigger
    else{
      for(j = 0; j < handler.length; j++){
        if(triggers[i].getHandlerFunction() === handler[j]){
         ScriptApp.deleteTrigger(triggers[i])
         found++
        }
      }
    }
  }
  // Output appeared after unsetting all kinds of this script's triggers
  if(trigger === 'unset'){
    if(found === 0){
      console.info('There is no trigger on set')
    }
    else{
      console.warn('Every forwarder\'s trigger has been unset')
    }
  }
}

// Postpone new execution if the execution is still running
function waitExecution(){
  // Verify if the execution is still running via Google Script Apps's API
  var apiGAS = UrlFetchApp.fetch('https://script.googleapis.com/v1/processes?userProcessFilter.functionName=forwarder&userProcessFilter.scriptId=' + EXEC_WAIT[1], {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}})
  // Filter the Google Script Apps's API response to match execution is running
  var responseGAS = JSON.parse(apiGAS).processes.filter(execution => {return execution.processType == 'TIME_DRIVEN' && execution.processStatus == 'RUNNING'})[0]
  var executionStatus = Object.values(responseGAS || {}).includes('RUNNING')
  // Ignore execution's duration less than the script's clock trigger
  try{
    var executionDuration = parseFloat(Object.values(responseGAS || {})[6].slice(0, -1)) > (CLOCK * 60)
  }
  // The execution is not running
  catch(supress){
    return false
  }
  // Postpone new execution according to Google Script Apps's API response result
  if(executionStatus && executionDuration){
    console.warn('POSTPONED: The execution is still running')
    return true
  }
  else{
    return false
  }
}

// Postpone the new trigger if there's a paging trigger exists
function waitPager(paging = false){
  var triggers = ScriptApp.getProjectTriggers()
  for(i = 0; i < triggers.length; i++){
    if(triggers[i].getHandlerFunction() === 'forwarderPager'){
      // Remove the paging trigger if it's disabled and delay the new trigger
      if(triggers[i].isDisabled()){
        unsetTrigger('forwarderPager')
        if(!paging){
          Utilities.sleep(1000 * 60 * DELAY)
        }
        return false
      }
      // Postpone the new trigger if the paging trigger is active
      else{
        console.warn('POSTPONED: The paging batch is still in progress')
        return true
      }
    }
  }
  return false
}

// Set a paging trigger with delay
function threadsPager(minute = DELAY, paged){
  // Re-validate the user's input
  if(preflightInspection(false, true)){
    // Set a paging trigger only when there's no other active paging trigger
    if(!waitPager(true)){
      ScriptApp
        .newTrigger('forwarderPager')
        .timeBased()
        .at(new Date((new Date()).getTime() + 1000 * 60 * minute))
        .create()
      console.warn('       PAGED: There are ' + paged + ' job(s) with maximum messages exceeded, and it will trigger a new batch in ' + minute + ' minute(s)')
    }
  }
}

// To differ the handler function of the main trigger
function forwarder(){
  gmailAutoForwarder(false, false)
}

// To differ the handler function of the paging trigger
function forwarderPager(){
  gmailAutoForwarder(true, false)
}

// Gmail Auto-forwarder
function gmailAutoForwarder(paging, trial){
  // If configured, wait for the running execution to end first
  if(EXEC_WAIT[0]){
    // Stop this execution at once if the execution is still running
    if(waitExecution()){
      return
    }
  }
  // Stop this execution at once if there's an active paging trigger
  if(!paging && waitPager()){
    return
  }
  else{
    var pages = []
    // Iterate the list of filters and recipients to perform forwarding
    for(i = 0; i < PASS.length; i++){
      var job = '[Job ' + (i + 1) + ']'
      var inInbox = GmailApp.search('in:inbox ' + PASS[i][0], 0, THREADS)
      console.info('   QUERY_FOR: ' + job + ' ' + PASS[i][0])
      // Skip this one filter process due to no matching message(s) found
      if(inInbox.length === 0){
        console.info('        IDLE: ' + job + ' There are no message(s) that match the query')
        pages.push(false)
        continue
      }
      // Proceed to forward the matching message(s)
      else{
        var messages = 0
        // Iterate the list of matching thread(s)
        for(j = 0; j < inInbox.length; j++){
          var messageThread = inInbox[j].getMessages()
          // Iterate the list of matching message(s)
          for(k = 0; k < messageThread.length; k++){
            if(!trial){
              // Gather the email components
              var sender = messageThread[k].getFrom()
              var subject = messageThread[k].getSubject().replace(/\s+/g,' ')
              var original = messageThread[k].getRawContent()
              var postscript = 'The original message is attached to keep its format intact.\n\nNOTE:\nIn case when viewing the attached message is clipped on its pop-up window.\nDownload and open it on an local email client (e.g., Thunderbird).'
              // Iterate the list of the matching message's recipient(s)
              for(l = 0; l < PASS[i][1].length; l++){
                // Debug to verify the job between the Inbox and Sent are in sync
                if(DEBUG){
                  console.warn('       DEBUG: ' + job + ' ### [Inbox] ### [Recipients No: ' + (l + 1) + '] ### [Forwarding & Deleting: \'' + subject.slice(0,25) + '...\']')
                }
                // Forward the matching message from Inbox to its recipient(s) as an attachment to keep its format intact
                if(ALIAS[0]){
                  GmailApp.sendEmail(PASS[i][1][l], subject, postscript, {from: GmailApp.getAliases()[ALIAS[1]], name: sender, replyTo: sender, attachments: [Utilities.newBlob(original, "message/rfc822", (subject + '.eml'))]})
                }
                else{
                  GmailApp.sendEmail(PASS[i][1][l], subject, postscript, {name: sender, replyTo: sender, attachments: [Utilities.newBlob(original, "message/rfc822", (subject + '.eml'))]})
                }
                Utilities.sleep(1000 * 60 * PAUSE)
              }
              // Delete the matching message permanently from Inbox
              do{
                try{
                  var retry = false
                  Gmail.Users.Messages.remove('me', messageThread[k].getId())
                }
                // Keep retrying to delete the matching message permanently if the execution request fails on the Google service
                catch(suppress){
                  var retry = true
                  if(DEBUG){
                    console.warn('       DEBUG: ' + job + ' ### [Inbox] ### [Total Recipients: ' + PASS[i][1].length + '] ### [       Retry Deleting: \'' + subject.slice(0,25) + '...\']')
                  }
                }
                Utilities.sleep(1000 * 60 * PAUSE)
              }
              while(retry)
              // Proceed to perform deletion in Sent
              var inSent = GmailApp.search('in:sent ' + PASS[i][0], 0, THREADS)
            }
            // Trial mode
            else{
              console.warn('DEBUG[TRIAL]: ' + job + ' ### [Inbox] ### [              Message: \'' + subject.slice(0,25) + '...\']')
            }
            // Count the matching message(s) per job
            messages++
          }
        }
        if(!trial){
          // Iterate the list of matching thread(s)
          for(j = 0; j < inSent.length; j++){
            var sentThread = inSent[j].getMessages()
            // Iterate the list of matching message(s)
            for(k = 0; k < sentThread.length; k++){
              var subject = sentThread[k].getSubject().replace(/\s+/g,' ')
              // Debug to verify the job between the Inbox and Sent are in sync
              if(DEBUG){
                console.warn('       DEBUG: ' + job + ' ### [ Sent] ### [Recipients No: ' + (k + 1) + '] ### [             Deleting: \'' + subject.slice(0,25) + '...\']')
              }
              // Delete the forwarded message from Sent permanently
              do{
                try{
                  var retry = false
                  Gmail.Users.Messages.remove('me', sentThread[k].getId())
                }
                // Keep retrying to delete the matching message permanently if the execution request fails on the Google service
                catch(suppress){
                  var retry = true
                  if(DEBUG){
                    console.warn('       DEBUG: ' + job + ' ### [ Sent] ### [Recipients No: ' + (k + 1) + '] ### [       Retry Deleting: \'' + subject.slice(0,25) + '...\']')
                  }
                }
                Utilities.sleep(1000 * 60 * PAUSE)
              }
              while(retry)
            }
          }
          console.info('FORWARDED_TO: ' + job + ' ' + PASS[(i)][1].join(', '))
          console.info('        DONE: ' + job + ' There are ' + messages + ' message(s) are forwarded then permanently deleted')
        }
        // Trial mode
        else{
          console.info('        DONE: ' + job + ' There are ' + messages + ' matched message(s)')
        }
        if(inInbox.length === THREADS){
          pages.push(inInbox.length)
        }
        else{
          pages.push(true)
        }
      }
    }
    // Set a paging trigger when the number of messages reaches the limit
    if(pages.includes(THREADS)){
      var paged = 0
      // Count every paged job
      for(i = 0; i < pages.length; i++){
        if(pages[i] === THREADS){
          paged++
        }
      }
      threadsPager(DELAY, paged)
    }
    else{
      // Remove the disabled paging trigger at the end of this paging batch
      if(paging){
        unsetTrigger('forwarderPager')
      }
      // Output appeared after every message was filtered
      if(pages.includes(true)){
        if(!trial){
          console.info('JOB_FINISHED: All messages are forwarded')
        }
        // Trial mode
        else{
          console.info('JOB_FINISHED: All messages are filtered. Run \'setTrigger\' to use the script')
        }
      }
      else{
        console.info('JOB_FINISHED: No messages matched the queries')
      }
    }
  }
}