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

//////////////////////////////////////////////////////////////////
// AUTHOR  : ADE DESTRIANTO                                     //
// TITLE   : GMAIL AUTO-FORWARDER                               //
// VERSION : 1.1 BUILD 20240317                                 //
// GITHUB  : https://github.com/destrianto/gmail_auto-forwarder //
//////////////////////////////////////////////////////////////////

// Set a clock trigger to this script
function setTrigger(){
  ScriptApp
    .newTrigger('forwarder')
    .timeBased()
    .everyMinutes(CLOCK)
    .create()
  console.info('The forwarder\'s clock trigger is adjusted every ' + CLOCK + ' minute(s)')
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
        console.warn('   POSTPONED: The paging batch is still in progress')
        return true
      }
    }
  }
  return false
}

// Set a paging trigger with delay
function threadsPager(minute = DELAY, paged){
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

// To differ the handler function of the paging trigger
function forwarderPager(){
  forwarder(true)
}

// Gmail Auto-forwarder
function forwarder(paging = false){
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
          var name = inbox[0].getFrom()
          // Forward the matching message from the Inbox and delete it permanently
          inbox[0].forward(PASS[i][1], {name: name, replyTo: name, subject: inbox[0].getSubject(), htmlBody: inbox[0].getBody()})
          Gmail.Users.Messages.remove('me', inbox[0].getId())
          // Debug to verify the job between the Inbox and Sent are in sequential sync
          if(DEBUG){
            console.warn('       DEBUG: ' + job + ' ### [Inbox] ### [Found: ' + (threads.length - j) + ' message(s)] ### [Forwarding & Deleting: 1 of ' + (threads.length - j) + ' message(s)] ### [Subject: \"' + inbox[0].getSubject() + '\"]')
          }
          Utilities.sleep(1000 * 60 * PAUSE)
          // Delete the forwarded message from Sent permanently
          var sent = GmailApp.search('in:sent ' + PASS[i][0], 0, THREADS)[0].getMessages()
          Gmail.Users.Messages.remove('me', sent[0].getId())
          // Debug to verify the job between the Inbox and Sent are in sequential sync
          if(DEBUG){
            console.warn('       DEBUG: ' + job + ' ### [ Sent] ### [Found: ' + sent.length + ' message(s)] ### [             Deleting: 1 of ' + sent.length + ' message(s)] ### [Subject: \"' + sent[0].getSubject() + '\"]')
          }
        }
        console.info('FORWARDED_TO: ' + job + ' ' + PASS[(i)][1].join(', '))
        console.info('        DONE: ' + job + ' There are ' + threads.length + ' message(s) are forwarded then permanently deleted')
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
        console.info('JOB_FINISHED: All messages forwarded')
      }
      else{
        console.info('JOB_FINISHED: No messages matched the queries')
      }
    }
  }
}
