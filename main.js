async function getCurrentBlock() {
  return await logseq.Editor.getCurrentBlock();
}

// Returns the page entity for the block; works for both MD and DB graphs.
async function getCurrentPage(block) {
  return await logseq.Editor.getPage(block.page.id);
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
 * Format datetime string using a custom format pattern
 * @param {string} timeString - Time string from Outlook (already in Eastern time)
 * @returns {string} - Formatted datetime string according to user configuration
 */
function formatTime(timeString) {
  // Parse the time string manually to avoid timezone conversion
  const date = new Date(timeString);
  
  // Extract components directly from the date object
  // Use UTC methods to avoid any local timezone adjustments
  const utcDate = new Date(date.getTime() + (date.getTimezoneOffset() * 60000));
  
  // Get the format pattern from settings
  const formatPattern = logseq.settings?.timeFormat || "h:mm a";
  
  // Extract date components
  const year = utcDate.getFullYear();
  const month = utcDate.getMonth();
  const day = utcDate.getDate();
  const hours24 = utcDate.getHours();
  const hours12 = hours24 % 12 || 12;
  const minutes = utcDate.getMinutes();
  const seconds = utcDate.getSeconds();
  const ampm = hours24 >= 12 ? 'PM' : 'AM';
  const ampmLower = ampm.toLowerCase();
  
  // Month names
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  const monthNamesShort = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
  ];
  
  // Day names
  const dayNames = [
    'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'
  ];
  const dayNamesShort = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  
  // Format tokens and their replacements
  const formatTokens = {
    // Year
    'YYYY': String(year),
    'YY': String(year).slice(-2),
    
    // Month
    'MMMM': monthNames[month],
    'MMM': monthNamesShort[month],
    'MM': String(month + 1).padStart(2, '0'),
    'M': String(month + 1),
    
    // Day of month
    'DD': String(day).padStart(2, '0'),
    'D': String(day),
    
    // Day of week
    'dddd': dayNames[utcDate.getDay()],
    'ddd': dayNamesShort[utcDate.getDay()],
    
    // Hour (24-hour)
    'HH': String(hours24).padStart(2, '0'),
    'H': String(hours24),
    
    // Hour (12-hour)
    'hh': String(hours12).padStart(2, '0'),
    'h': String(hours12),
    
    // Minutes
    'mm': String(minutes).padStart(2, '0'),
    'm': String(minutes),
    
    // Seconds
    'ss': String(seconds).padStart(2, '0'),
    's': String(seconds),
    
    // AM/PM
    'A': ampm,
    'a': ampmLower
  };
  
  // --- START OF FIX ---
  // Create a single regex that matches any of the tokens, longest first.
  const sortedTokens = Object.keys(formatTokens).sort((a, b) => b.length - a.length);
  const regex = new RegExp(sortedTokens.join('|'), 'g');

  // Use the .replace() method with a replacer function.
  // This finds all tokens in one pass and replaces them with their corresponding values.
  return formatPattern.replace(regex, (match) => formatTokens[match]);
  // --- END OF FIX ---
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



// Returns filtered attendee names as an array (node property accepts names directly).
function filterAttendees(attendees) {
  if (!attendees || attendees.length === 0) return [];
  const excludeName = logseq.settings?.excludeUserName || "";
  return excludeName
    ? attendees.filter(name => name !== excludeName)
    : attendees;
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

// Set a DB property on a block; skips empty values unless includeEmptyFields is on.
async function setEventProperty(blockUUID, key, value) {
  const includeEmpty = logseq.settings?.includeEmptyFields || false;
  if (!value && !includeEmpty) return;
  await logseq.Editor.upsertBlockProperty(blockUUID, key, value ?? "");
}

async function getEvents(e) {
  try {
    const currentBlock = await getCurrentBlock();
    if (!currentBlock) {
      logseq.UI.showMsg('Place the cursor in a block before running /Get Events', 'warning');
      return;
    }

    const page = await getCurrentPage(currentBlock);

    // DB graph journals expose journalDay; fall back to journal? flag check.
    const journalDay = page?.journalDay ?? page?.["journal-day"];
    const isJournal = !!journalDay || page?.["journal?"] === true || page?.journal === true;
    if (!isJournal) {
      logseq.UI.showMsg('Get Events only works on journal pages', 'warning');
      return;
    }

    const pageDate = journalDayToDate(journalDay);
    const apiDateString = formatDateForAPI(pageDate);
    console.log('Fetching events for:', apiDateString);

    const events = await fetchEventsFromAPI(apiDateString);
    if (events.length === 0) {
      logseq.UI.showMsg(`No events found for ${apiDateString}`, 'warning');
      return;
    }

    const sortedEvents = events.sort((a, b) => new Date(a.start) - new Date(b.start));

    // Insert in reverse order with before:true so final order is chronological.
    for (let i = sortedEvents.length - 1; i >= 0; i--) {
      const event = sortedEvents[i];
      const subject = formatEventSubject(event.subject, event.isRecurring, event.meetingLinks);

      const insertedBlock = await logseq.Editor.insertBlock(
        e.uuid,
        subject,
        { sibling: true, before: true }
      );

      if (!insertedBlock?.uuid) continue;
      const uuid = insertedBlock.uuid;

      await setEventProperty(uuid, 'event-time', formatTime(event.start));
      await setEventProperty(uuid, 'event-duration', calculateDuration(event.start, event.end));
      await setEventProperty(uuid, 'event-location', event.location || "");
      await setEventProperty(uuid, 'recurring', event.isRecurring ? true : null);

      const attendees = filterAttendees(event.attendees);
      if (attendees.length > 0) {
        await logseq.Editor.upsertBlockProperty(uuid, 'attendees', attendees);
      }

      const desc = formatDescription(event.description || "");
      await setEventProperty(uuid, 'event-description', desc);
    }

  } catch (error) {
    console.error('Error in getEvents:', error);
    logseq.UI.showMsg(`Get Events error: ${error.message}`, 'error');
  }
}

// Registers the typed DB properties this plugin writes; safe to call on every load.
async function registerProperties() {
  const upsert = logseq.Editor.upsertProperty;
  await upsert('event-time',        { type: 'default' });
  await upsert('event-duration',     { type: 'default' });
  await upsert('event-location',     { type: 'default' });
  await upsert('event-description',  { type: 'default' });
  await upsert('recurring',          { type: 'checkbox' });
  await upsert('attendees',          { type: 'node', cardinality: 'many' });
}

const main = async () => {
  console.log('Get Outlook Events Plugin Loaded');

  if (typeof logseq.Editor.upsertBlockProperty !== 'function' ||
      typeof logseq.Editor.upsertProperty !== 'function') {
    logseq.UI.showMsg(
      'Get Outlook Events requires a Logseq DB build. Please update Logseq.',
      'error'
    );
    return;
  }

  await registerProperties();

  logseq.useSettingsSchema([
    {
      key: 'excludeUserName',
      type: 'string',
      default: '',
      title: 'Excluded User Name',
      description: "Your name as returned by Outlook — excluded from the attendees property (e.g. 'Simpson, Homer')"
    },
    {
      key: 'apiUrl',
      type: 'string',
      default: 'http://localhost:5000',
      title: 'API URL',
      description: 'URL of the Outlook Events API service'
    },
    {
      key: 'bracketEvents',
      type: 'enum',
      default: 'none',
      title: 'Event Title as Node',
      description: 'Make event titles page-linked nodes in the block content',
      enumChoices: ['all', 'recurring', 'none'],
      enumPicker: 'select'
    },
    {
      key: 'timeFormat',
      type: 'string',
      default: 'h:mm a',
      title: 'Time Format',
      description: 'Tokens: h/hh/H/HH (hour), mm (min), a/A (am/pm). Examples: h:mm a → 9:30 AM  |  HH:mm → 09:30'
    },
    {
      key: 'meetingBaseUrls',
      type: 'string',
      default: 'https://teams.microsoft.com,https://zoom.us,https://meet.google.com',
      title: 'Meeting Base URLs',
      description: 'Comma-separated base URLs to detect meeting links in event descriptions'
    },
    {
      key: 'includeEmptyFields',
      type: 'boolean',
      default: false,
      title: 'Include Empty Fields',
      description: 'Write properties even when the value is empty (e.g. event-location with no location set)'
    },
    {
      key: 'descriptionMaxLength',
      type: 'number',
      default: 200,
      title: 'Description Max Length',
      description: 'Max characters from event descriptions (0 = no limit)'
    }
  ]);

  logseq.Editor.registerSlashCommand('Get Events', getEvents);
};

logseq.ready(main).catch(console.error);