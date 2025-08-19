/**
 * Current Block
 * @returns {object} - An object representing the block the plug-in is triggered in.
 */
async function getCurrentBlock() {
  let currentBlock = await logseq.App.getCurrentBlock();
  if (currentBlock) {
    currentBlockId = currentBlock.uuid; // Store the block ID (UUID) in the variable
    console.log("Current block ID:", currentBlockId);
  } else {
    console.log("No block is currently selected.");
  }
  return currentBlock;
}

/**
 * @param {object} block - An object representing the block the plug-in is triggered in.
 * @returns - An object representing the page that hosts the provided block object.
 */
async function getCurrentPage(block) {
  const page = await logseq.Editor.getPage(block.page.id);
  if (page) {
    currentPageName = page.name;
    console.log("Page name:", currentPageName);
    
    if (page.journal) {
      console.log("This is a journal page. Date:", page.journal["date"]);
    } else {
      console.log("This is not a journal page.");
    }
  } else {
    console.log("No page is currently open.");
  }
  
  return page;
}

/**
 * @param {string} journalDay - The date of the Journal Page.
 * @returns {Date} - The date converted from the provided string.
 */
function journalDayToDate(journalDay) {
  const y = Math.floor(journalDay / 10000);
  const m = Math.floor((journalDay % 10000) / 100) - 1;
  const d = journalDay % 100;
  return new Date(y, m, d);
}

/**
 * Format a date string to YYYY-MM-DD format for the API
 * @param {Date} date - The date to format
 * @returns {string} - Formatted date string
 */
function formatDateForAPI(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Format time string to 12-hour format (HH:MM AM/PM) without timezone conversion
 * @param {string} timeString - Time string from Outlook (already in Eastern time)
 * @returns {string} - Formatted time string
 */
function formatTime(timeString) {
  // Parse the time string manually to avoid timezone conversion
  const date = new Date(timeString);
  
  // Extract hours and minutes directly from the date object
  // Use UTC methods to avoid any local timezone adjustments
  const utcDate = new Date(date.getTime() + (date.getTimezoneOffset() * 60000));
  let hours = utcDate.getHours();
  const minutes = utcDate.getMinutes();
  
  // Convert to 12-hour format
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12; // the hour '0' should be '12'
  
  // Format minutes with leading zero if needed
  const minutesStr = minutes < 10 ? '0' + minutes : minutes;
  
  return `${hours}:${minutesStr} ${ampm}`;
}

/**
 * Calculate duration between two times as a time quantity
 * @param {string} startTime - Start time string (already in Eastern time)
 * @param {string} endTime - End time string (already in Eastern time)
 * @returns {string} - Duration in format "HH:MM:SS"
 */
function calculateDuration(startTime, endTime) {
  const start = new Date(startTime);
  const end = new Date(endTime);
  
  // Calculate difference in milliseconds
  const diffMs = end.getTime() - start.getTime();
  
  // Convert to hours, minutes, seconds
  const totalSeconds = Math.floor(diffMs / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  
  // Format as HH:MM:SS
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

/**
 * Format attendees list with square brackets
 * @param {Array} attendees - Array of attendee names
 * @returns {string} - Comma-separated list with each name in square brackets
 */
function formatAttendees(attendees) {
  if (!attendees || attendees.length === 0) {
    return "";
  }
  return attendees.map(name => `[[${name}]]`).join(", ");
}

/**
 * Fetch events from the Outlook API
 * @param {string} dateString - Date in YYYY-MM-DD format
 * @returns {Promise<Array>} - Array of events or empty array if error
 */
async function fetchEventsFromAPI(dateString) {
  try {
    const response = await fetch(`http://localhost:5000/events/${dateString}`);
    
    if (!response.ok) {
      console.error(`API request failed: ${response.status} ${response.statusText}`);
      return [];
    }
    
    const data = await response.json();
    
    if (data.success) {
      console.log(`Found ${data.events.length} events for ${dateString}`);
      return data.events;
    } else {
      console.error('API returned error:', data.error);
      return [];
    }
  } catch (error) {
    console.error('Error fetching events from API:', error);
    return [];
  }
}


// Insert the day's list of events from the local Outlook calendar.
// Insert the day's list of events from the local Outlook calendar.
async function getEvents(e) {
  console.log('=== getEvents function called ===');
  console.log('Trigger block UUID:', e.uuid);
  
  try {
    const currentBlock = await getCurrentBlock();
    const currentPage = await getCurrentPage(currentBlock);
    
    if (!currentPage?.journalDay) {
      console.log("Not on a journal page");
      logseq.Editor.insertBlock(e.uuid, `Error: This command only works on journal pages`, {before: true});
      return;
    }
    
    const pageDate = journalDayToDate(currentPage.journalDay);
    const apiDateString = formatDateForAPI(pageDate);
    
    console.log("Fetching events for date:", apiDateString);
    
    // Fetch events from API
    const events = await fetchEventsFromAPI(apiDateString);
    
    if (events.length === 0) {
      logseq.Editor.insertBlock(e.uuid, `No events found for ${apiDateString}`, {before: true});
      return;
    }
    
    console.log(`Processing ${events.length} events`);
    
    // Sort events by start time (earliest to latest)
    const sortedEvents = events.sort((a, b) => {
      const startTimeA = new Date(a.start);
      const startTimeB = new Date(b.start);
      return startTimeA - startTimeB; // This gives us earliest to latest
    });
    
    console.log('Events sorted by start time');
    
    // Insert each event as a block
    for (let i = 0; i < sortedEvents.length; i++) {
      const event = sortedEvents[i];
      console.log(`Processing event ${i + 1}:`, event.subject);
      
      // Format the complete event block content
      const eventContent = [
        event.subject,
        `event-time:: ${formatTime(event.start)}`,
        `event-duration:: ${calculateDuration(event.start, event.end)}`,
        `attendees:: ${formatAttendees(event.attendees)}`
      ].join('\n');
      
      // Insert each event as a sibling after the trigger block
      const insertedBlock = await logseq.Editor.insertBlock(
        e.uuid, 
        eventContent, 
        { sibling: true, before: true }
      );
      
      console.log('Event block inserted, UUID:', insertedBlock?.uuid);
      console.log(`Inserted event: ${event.subject}`);
    }
    
    console.log('=== getEvents function completed ===');
    
  } catch (error) {
    console.error('Error in getEvents:', error);
    logseq.Editor.insertBlock(e.uuid, `Error fetching events: ${error.message}`, {before: true});
  }
}

// The main app
const main = async () => {
  console.log('Get Outlook Events Plugin Loaded');
  
  logseq.Editor.registerSlashCommand('Get Events', async (e) => {
    getEvents(e);
  });
}

logseq.ready(main).catch(console.error);