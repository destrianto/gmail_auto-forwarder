///////////////////////////////////////////////////////////////////////
// AVAILABLE FUNCTIONS FOR THE USER TO RUN                           //
// • 'preflightInspection': Validate the user's input manually.      //
// • 'setTrigger': Install a trigger for this script.                //
// • 'unsetTrigger': Uninstall all trigger kinds of this script.     //
// DO NOT RUN THESE FUNCTIONS MANUALLY. RUN THEM AT YOUR OWN RISK!!! //
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

//////////////////////////////////////////////////////////////////
// AUTHOR  : ADE DESTRIANTO                                     //
// TITLE   : GMAIL AUTO-FORWARDER                               //
// VERSION : 1.2 BUILD 20240329                                 //
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
  for(var i = 0; i < triggers.length; i++){
    // Unset specified trigger
    if(trigger !== 'unset'){
      if(triggers[i].getHandlerFunction() === trigger){
        ScriptApp.deleteTrigger(triggers[i])
      }
    }
    // Unset all kinds of this script's trigger
    else{
      for(var j = 0; j < handler.length; j++){
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

// Postpone the new trigger if there's a paging trigger exists
function waitPager(paging = false){
  var triggers = ScriptApp.getProjectTriggers()
  for(var i = 0; i < triggers.length; i++){
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
  // Stop this execution at once if there's an active paging trigger
  if(!paging && waitPager()){
    return
  }
  else{
    var count = []
    // Iterate the list of filters and recipients to perform forwarding
    for(var i = 0; i < PASS.length; i++){
      var job = '[Job ' + (i + 1) + ']'
      var threads = GmailApp.search(PASS[i][0], 0, THREADS)
      console.info('   QUERY_FOR: ' + job + ' ' + PASS[i][0])
      // Skip this one filter process due to no matching message(s) found
      if(threads.length === 0){
        console.info('        IDLE: ' + job + ' There are no message(s) that match the query')
        count.push(false)
        continue
      }
      // Proceed to forward the matching message(s)
      else{
        // Iterate the list of matching message(s)
        for(var j = 0; j < threads.length; j++){
          var inbox = threads[j].getMessages()
          if(!trial){
            var name = inbox[0].getFrom()
            // Forward the matching message from the Inbox and delete it permanently
            inbox[0].forward(PASS[i][1], {name: name, replyTo: name, subject: inbox[0].getSubject(), htmlBody: inbox[0].getBody()})
            Gmail.Users.Messages.remove('me', inbox[0].getId())
            // Debug to verify the job between the Inbox and Sent are in sequential sync
            if(DEBUG){
              console.warn('       DEBUG: ' + job + ' ### [Inbox] ### [Found: ' + (threads.length - j) + ' message(s)] ### [Forwarding & Deleting: 1 of ' + (threads.length - j) + ' message(s)] ### [Subject: \'' + inbox[0].getSubject() + '\']')
            }
            Utilities.sleep(1000 * 60 * PAUSE)
            // Delete the forwarded message from Sent permanently
            var sent = GmailApp.search('in:sent ' + PASS[i][0], 0, THREADS)[0].getMessages()
            Gmail.Users.Messages.remove('me', sent[0].getId())
            // Debug to verify the job between the Inbox and Sent are in sequential sync
            if(DEBUG){
              console.warn('       DEBUG: ' + job + ' ### [ Sent] ### [Found: ' + sent.length + ' message(s)] ### [             Deleting: 1 of ' + sent.length + ' message(s)] ### [Subject: \'' + sent[0].getSubject() + '\']')
            }
          }
          // Trial mode
          else{
            console.warn('DEBUG[TRIAL]: ' + job + ' ### [Inbox] ### [Found: ' + threads.length + ' message(s)] ### [Message: ' + (j + 1) + ' of ' + threads.length + '] ### [Subject: \'' + inbox[0].getSubject() + '\']')
          }
        }
        if(!trial){
          console.info('FORWARDED_TO: ' + job + ' ' + PASS[(i)][1].join(', '))
          console.info('        DONE: ' + job + ' There are ' + threads.length + ' message(s) are forwarded then permanently deleted')
        }
        // Trial mode
        else{
          console.info('        DONE: ' + job + ' There are ' + threads.length + ' matched message(s)')
        }
        if(threads.length === THREADS){
          count.push(threads.length)
        }
        else{
          count.push(true)
        }
      }
    }
    // Set a paging trigger when the number of messages reaches the limit
    if(count.includes(THREADS)){
      var paged = 0
      // Count every paged job
      for(var i = 0; i < count.length; i++){
        if(count[i] === THREADS){
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
      if(count.includes(true)){
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