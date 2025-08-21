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
 * Format time string to 12-hour or 24-hour format without timezone conversion
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
  
  // Format minutes with leading zero if needed
  const minutesStr = minutes < 10 ? '0' + minutes : minutes;
  
  // Get time format setting (default to 12-hour)
  const timeFormat = logseq.settings?.timeFormat || "12";
  
  if (timeFormat === "24") {
    // 24-hour format
    const hoursStr = hours < 10 ? '0' + hours : hours;
    return `${hoursStr}:${minutesStr}`;
  } else {
    // 12-hour format
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // the hour '0' should be '12'
    return `${hours}:${minutesStr} ${ampm}`;
  }
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
 * Create meeting links for an event
 * @param {Array} meetingLinks - Array of meeting link URLs
 * @returns {string} - Markdown links string
 */
function createMeetingLinks(meetingLinks) {
  if (!meetingLinks || meetingLinks.length === 0) {
    return "";
  }
  
  let linksHtml = "";
  
  meetingLinks.forEach((link, index) => {
    // Use "Join Meeting" for all links, with numbering for multiple links
    let linkText = meetingLinks.length > 1 ? `Join Meeting ${index + 1}` : "Join Meeting";
    
    // Create the markdown link
    linksHtml += `[${linkText}](${link})`;
    
    // Add space between multiple links
    if (index < meetingLinks.length - 1) {
      linksHtml += " ";
    }
  });
  
  return " " + linksHtml; // Add leading space to separate from subject/emoji
}

/**
 * Format event subject based on bracket settings and add recurring emoji and meeting links
 * @param {string} subject - The event subject
 * @param {boolean} isRecurring - Whether the event is recurring
 * @param {Array} meetingLinks - Array of meeting link URLs
 * @returns {string} - Formatted subject with brackets, recurring emoji, and meeting links
 */
function formatEventSubject(subject, isRecurring = false, meetingLinks = []) {
  const bracketSetting = logseq.settings?.bracketEvents || "none";
  
  let formattedSubject = subject;
  
  // First apply brackets based on setting
  switch (bracketSetting) {
    case "all":
      formattedSubject = `[[${subject}]]`;
      break;
    case "recurring":
      formattedSubject = isRecurring ? `[[${subject}]]` : subject;
      break;
    case "none":
    default:
      formattedSubject = subject;
      break;
  }
  
  // Then add recurring emoji outside the brackets if it's a recurring event
  if (isRecurring) {
    formattedSubject = `${formattedSubject} 🔃`;
  }
  
  // Finally add meeting links after subject and emoji
  const meetingLinks_formatted = createMeetingLinks(meetingLinks);
  formattedSubject = `${formattedSubject}${meetingLinks_formatted}`;
  
  return formattedSubject;
}

/**
 * Format event description by truncating if needed
 * @param {string} description - The event description
 * @returns {string} - Formatted/truncated description
 */
function formatDescription(description) {
  if (!description) return "";
  
  const maxLength = logseq.settings?.descriptionMaxLength || 0;
  
  if (maxLength > 0 && description.length > maxLength) {
    return description.substring(0, maxLength).trim() + "...";
  }
  
  return description.trim();
}

/**
 * Format event block content based on user template and handle child blocks
 * @param {object} event - The event object from API
 * @returns {object} - Object with mainContent and childBlocks array
 */
function formatEventContent(event) {
  const template = logseq.settings?.outputFormat || 
    "{subject}\\nevent-time:: {time}\\nevent-duration:: {duration}\\nattendees:: {attendees}";
  const includeEmpty = logseq.settings?.includeEmptyFields || false;
  
  // Format the event subject with brackets, emoji, and meeting links
  const formattedSubject = formatEventSubject(event.subject, event.isRecurring, event.meetingLinks);
  
  // Prepare all possible variables
  const variables = {
    subject: formattedSubject,
    time: formatTime(event.start),
    duration: calculateDuration(event.start, event.end),
    attendees: formatAttendees(event.attendees),
    location: event.location || "",
    description: formatDescription(event.description || "")
  };
  
  // Replace variables in template
  let content = template;
  
  // Handle each variable replacement
  Object.keys(variables).forEach(key => {
    const value = variables[key];
    const placeholder = `{${key}}`;
    
    if (content.includes(placeholder)) {
      if (!includeEmpty && !value) {
        // Remove the entire line if the field is empty and includeEmpty is false
        const lines = content.split('\\n');
        content = lines.filter(line => {
          if (line.includes(placeholder)) {
            // Only remove lines that are just property assignments (contain ::)
            return line.includes('::') ? false : true;
          }
          return true;
        }).join('\\n');
      } else {
        // Replace the placeholder with the value
        content = content.replace(new RegExp(`\\{${key}\\}`, 'g'), value);
      }
    }
  });
  
  // Convert \\n to actual newlines
  content = content.replace(/\\n/g, '\n');
  
  // Split content by ---CHILD--- delimiter to separate main content from child blocks
  const parts = content.split('---CHILD---');
  const mainContent = parts[0].trim();
  const childBlocks = parts.slice(1).map(block => block.trim()).filter(block => block.length > 0);
  
  return {
    mainContent,
    childBlocks
  };
}

/**
 * @param {Array} attendees - Array of attendee names
 * @returns {string} - Comma-separated list with each name in square brackets
 */
function formatAttendees(attendees) {
  if (!attendees || attendees.length === 0) {
    return "";
  }
  
  // Get the user's configured name to exclude
  const excludeName = logseq.settings?.excludeUserName || "";
  
  // Filter out the user's name if configured
  const filteredAttendees = excludeName 
    ? attendees.filter(name => name !== excludeName)
    : attendees;
  
  return filteredAttendees.map(name => `[${name}]`).join(", ");
}

/**
 * Fetch events from the Outlook API
 * @param {string} dateString - Date in YYYY-MM-DD format
 * @returns {Promise<Array>} - Array of events or empty array if error
 */
async function fetchEventsFromAPI(dateString) {
  try {
    // Get the configured API URL, default to localhost:5000
    const apiUrl = logseq.settings?.apiUrl || 'http://localhost:5000';
    const meetingBaseUrls = logseq.settings?.meetingBaseUrls || '';
    
    // Build the URL with meeting URLs parameter if configured
    let url = `${apiUrl}/events/${dateString}`;
    if (meetingBaseUrls.trim()) {
      url += `?meeting_urls=${encodeURIComponent(meetingBaseUrls)}`;
    }
    
    const response = await fetch(url);
    
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
    
    // Insert in reverse order with before: true to get chronological order
    for (let i = sortedEvents.length - 1; i >= 0; i--) {
      const event = sortedEvents[i];
      console.log(`Processing event ${i + 1}:`, event.subject);
      
      // Format the complete event block content using the user's template
      const eventData = formatEventContent(event);
      
      // Insert the main event block
      const insertedBlock = await logseq.Editor.insertBlock(
        e.uuid, 
        eventData.mainContent, 
        { sibling: true, before: true }
      );
      
      console.log('Event block inserted, UUID:', insertedBlock?.uuid);
      
      // Insert any child blocks
      if (eventData.childBlocks.length > 0 && insertedBlock?.uuid) {
        for (const childContent of eventData.childBlocks) {
          const childBlock = await logseq.Editor.insertBlock(
            insertedBlock.uuid,
            childContent,
            { sibling: false }  // Insert as child, not sibling
          );
          console.log('Child block inserted, UUID:', childBlock?.uuid);
        }
      }
      
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
  
  // Register plugin settings
  logseq.useSettingsSchema([
    {
      key: "excludeUserName",
      type: "string",
      default: "",
      title: "Excluded User Name",
      description: "Enter your name as it is returned by Outlook to exclude it from attendees lists (e.g., 'Simpson, Homer' or 'Diana Prince')"
    },
    {
      key: "apiUrl",
      type: "string", 
      default: "http://localhost:5000",
      title: "API URL",
      description: "The URL of the Outlook Events API service"
    },
    {
      key: "bracketEvents",
      type: "enum",
      default: "none",
      title: "Add Double Brackets to Event Titles",
      description: "Choose when to add [[double brackets]] around event subjects to create Logseq page links\n\nOptions:\n- all: Add brackets to all event titles\n- recurring: Only add brackets to recurring events\n- none: No brackets (default)",
      enumChoices: ["all", "recurring", "none"],
      enumPicker: "select"
    },
    {
      key: "timeFormat",
      type: "enum",
      default: "12",
      title: "Time Format",
      description: "Choose between 12-hour (9:00 AM) or 24-hour (09:00) time format",
      enumChoices: ["12", "24"],
      enumPicker: "select"
    },
    {
      key: "meetingBaseUrls",
      type: "string",
      default: "https://teams.microsoft.com,https://zoom.us,https://meet.google.com",
      title: "Meeting Base URLs",
      description: "Comma-separated list of base URLs to look for meeting links (e.g., 'https://teams.microsoft.com,https://zoom.us,https://meet.google.com')"
    },
    {
      key: "outputFormat",
      type: "string",
      inputAs: "textarea",
      default: "{subject}\\nevent-time:: {time}\\nevent-duration:: {duration}\\nattendees:: {attendees}",
      title: "Output Format Template",
      description: "Customize the format of event blocks. Available variables:\n\n{subject}, {time}, {duration}, {attendees}, {location}, {description}\n\nAdd ---CHILD--- to start a new child blocks."
    },
    {
      key: "includeEmptyFields",
      type: "boolean",
      default: false,
      title: "Include Empty Fields",
      description: "Whether to include fields in the output even when they are empty (e.g., show 'location::' even if no location is set)"
    },
    {
      key: "descriptionMaxLength",
      type: "number",
      default: 200,
      title: "Description Max Length",
      description: "Maximum number of characters to include from event descriptions (0 = no limit)"
    }
  ]);
  
  logseq.Editor.registerSlashCommand('Get Events', async (e) => {
    getEvents(e);
  });
}

logseq.ready(main).catch(console.error);