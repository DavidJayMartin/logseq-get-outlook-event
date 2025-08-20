# Outlook Events Logseq Plugin

A Logseq plugin that fetches events from your local Outlook calendar and inserts them into your journal pages. This plugin consists of two components: a Logseq plugin (JavaScript) and a local API service (Python) that reads from your Outlook OST file.

## Features

- 📅 Fetches events from your local Outlook calendar
- 📝 Inserts events directly into Logseq journal pages
- ⏰ Displays events in chronological order (earliest to latest)
- 👥 Shows attendees for each event (with configurable exclusions)
- ⏱️ Calculates and displays event duration
- 🏷️ Uses Logseq properties for structured data
- 🔃 Shows recurring event indicator
- 🔗 **NEW**: Automatically detects and creates meeting links
- ⚙️ **NEW**: Fully configurable through Logseq plugin settings
- 📋 **NEW**: Customizable output format templates

## Components

### 1. Logseq Plugin (`main.js`)
The main plugin that runs within Logseq and provides the `/Get Events` slash command.

### 2. API Service (`outlook_events_api.py`)
A local Flask API service that connects to your Outlook installation and retrieves calendar events.

## Prerequisites

- **Logseq** installed and running
- **Microsoft Outlook** installed on Windows with access to your calendar
- **Python 3.x** installed
- Required Python packages: `flask`, `pywin32`

## Installation

### Step 1: Set up the Python API Service

1. Save the `outlook_events_api.py` file to a directory of your choice

2. Install required Python packages:
   ```bash
   pip install flask pywin32
   ```

3. Start the API service:
   ```bash
   python outlook_events_api.py
   ```
   
   The service will start on `http://localhost:5000`

### Step 2: Install the Logseq Plugin

1. Create a new folder in your Logseq plugins directory (usually `~/.logseq/plugins/`)
2. Name the folder something like `outlook-events-plugin`
3. Copy the `main.js` file into this folder
4. Create a `package.json` file in the same folder:
   ```json
   {
     "name": "outlook-events-plugin",
     "version": "1.0.0",
     "description": "Fetch Outlook calendar events into Logseq",
     "main": "main.js",
     "logseq": {
       "id": "outlook-events-plugin",
       "title": "Outlook Events Plugin"
     }
   }
   ```
5. Restart Logseq or reload plugins

## Usage

### Starting the Services

1. **Start the Python API service first:**
   ```bash
   python outlook_events_api.py
   ```
   Keep this running while using the plugin.

2. **Open Logseq** and navigate to a journal page (daily note)

3. **Use the slash command:**
   - Type `/Get Events` in any block
   - The plugin will fetch events for that journal date
   - Events will be inserted as blocks below your cursor

### Event Format

Each event is inserted as a block with the following default format:

```
Event Title 🔃 [Join Meeting](https://teams.microsoft.com/l/meetup-join/...)
event-time:: 9:00 AM
event-duration:: 01:00:00
attendees:: [John Doe], [Jane Smith]
```

Where:
- 🔃 indicates a recurring event
- Meeting links are automatically detected and added as clickable links
- All fields are customizable through plugin settings

## Configuration

The plugin is now fully configurable through Logseq's plugin settings. Access settings via:
**Settings → Plugins → Outlook Events Plugin**

### Available Settings

#### **Exclude User Name**
- **Description**: Enter your name to exclude it from attendees lists
- **Default**: Empty (no exclusion)
- **Example**: `"Smith, John"` or `"Jane Doe"`

#### **API URL**
- **Description**: URL of the Outlook Events API service
- **Default**: `http://localhost:5000`
- **Usage**: Change if running the API on a different port or machine

#### **Add Double Brackets to Event Titles**
- **Options**: 
  - `all` - Add brackets to all event titles
  - `recurring` - Only add brackets to recurring events
  - `none` - No brackets (default)
- **Purpose**: Creates Logseq page links for events

#### **Time Format**
- **Options**: 
  - `12` - 12-hour format with AM/PM (e.g., 9:00 AM, 2:30 PM)
  - `24` - 24-hour format (e.g., 09:00, 14:30)
- **Default**: `12` (12-hour format)
- **Purpose**: Controls how event times are displayed

#### **Meeting Base URLs**
- **Description**: Comma-separated list of base URLs to search for meeting links
- **Default**: `https://teams.microsoft.com,https://zoom.us,https://meet.google.com`
- **Example**: `https://teams.microsoft.com,https://zoom.us,https://webex.com`
- **Note**: The plugin will find any URLs in event descriptions that start with these base URLs

#### **Output Format Template**
- **Description**: Customize how events are displayed using template variables
- **Default**: `{subject}\nevent-time:: {time}\nevent-duration:: {duration}\nattendees:: {attendees}`
- **Available Variables**:
  - `{subject}` - Event title (with brackets, emoji, and meeting links)
  - `{time}` - Start time in 12-hour format
  - `{duration}` - Event duration in HH:MM:SS format
  - `{attendees}` - Formatted attendee list
  - `{location}` - Event location
  - `{description}` - Event description (truncated if configured)
- **Special Features**:
  - Use `\n` for line breaks
  - Use `---CHILD---` to create child blocks
- **Example Custom Template**:
  ```
  ## {subject}
  **Time:** {time} ({duration})
  **Location:** {location}
  ---CHILD---
  **Attendees:** {attendees}
  ---CHILD---
  **Notes:** {description}
  ```

#### **Include Empty Fields**
- **Description**: Whether to show fields even when they're empty
- **Default**: `false`
- **Example**: When `true`, shows `location::` even if no location is set

#### **Description Max Length**
- **Description**: Maximum characters to include from event descriptions
- **Default**: `200`
- **Usage**: Set to `0` for no limit

### Meeting Links Feature

The plugin automatically scans event descriptions for meeting links and adds clickable "Join Meeting" links to event titles. 

**How it works:**
1. Configure the meeting base URLs in settings (e.g., `https://teams.microsoft.com,https://zoom.us`)
2. The plugin searches event descriptions for URLs starting with these base URLs
3. Found links are added as `[Join Meeting](URL)` next to the event title
4. Multiple links are numbered: `[Join Meeting 1](URL1) [Join Meeting 2](URL2)`

### API Endpoints

The Python service provides several endpoints:

- `GET /health` - Health check
- `GET /events/YYYY-MM-DD` - Get events for specific date
- `GET /events/today` - Get today's events
- `GET /events?date=YYYY-MM-DD&meeting_urls=url1,url2` - Get events with meeting URL detection

Examples:
```
http://localhost:5000/events/2025-08-19
http://localhost:5000/events/today
http://localhost:5000/events/2025-08-19?meeting_urls=https://teams.microsoft.com,https://zoom.us
```

## Advanced Customization

### Custom Output Templates

You can create complex event layouts using the output format template. Here are some examples:

#### Minimal Format
```
{subject}
{time}
```

#### Detailed Format with Child Blocks
```
{subject}
event-time:: {time}
event-duration:: {duration}
---CHILD---
location:: {location}
---CHILD---
description:: {description}
---CHILD---
attendees:: {attendees}
```

#### Meeting-Focused Format
```
## {subject}
⏰ {time} | ⏱️ {duration}
📍 {location}
👥 {attendees}
```

### Recurring Events

Recurring events are automatically identified and marked with a 🔃 emoji. You can configure bracket settings to automatically create page links for recurring events only.

## Troubleshooting

### Common Issues

1. **"No events found" message:**
   - Ensure the Python API service is running
   - Check that you're on a journal page (daily note)
   - Verify events exist in Outlook for that date

2. **API connection errors:**
   - Confirm the Python service is running on localhost:5000
   - Check the API URL setting in plugin configuration
   - Ensure Windows Firewall isn't blocking the connection

3. **Meeting links not appearing:**
   - Verify meeting base URLs are configured correctly
   - Check that meeting links in Outlook match the configured base URLs
   - Ensure the Python service is receiving the meeting URLs parameter

4. **Plugin not loading:**
   - Verify the plugin folder structure
   - Check Logseq's plugin settings
   - Review browser console for JavaScript errors

5. **Formatting issues:**
   - Check the output format template for syntax errors
   - Ensure `\n` is used for line breaks, not actual newlines
   - Verify template variable names are spelled correctly

### Debug Information

The plugin logs detailed information to the browser console. Open Developer Tools (F12) to view debug output including:
- API requests and responses
- Event processing details
- Template rendering information

## Technical Details

### Time Handling

- Events are fetched in Eastern Time (Outlook's local timezone)
- Times are displayed in either 12-hour format (9:00 AM) or 24-hour format (09:00) based on plugin settings
- Duration is always calculated and displayed as HH:MM:SS format

### Data Flow

1. User triggers `/Get Events` command on a journal page
2. Plugin extracts the journal date and configuration settings
3. API request sent to Python service with meeting URL parameters
4. Python service queries Outlook COM interface and extracts meeting links
5. Events returned as JSON with meeting links included
6. Plugin formats events using configured templates and inserts as Logseq blocks

### Meeting Link Detection

The Python service uses regular expressions to find URLs in event descriptions that start with the configured base URLs. It handles:
- Multiple meeting links per event
- Various URL formats and parameters
- Cleaning up URLs (removing trailing punctuation)

## Security Notes

- The API service only runs locally (localhost)
- No external network connections required
- Calendar data never leaves your local machine
- Uses Windows COM interface for secure Outlook access
- Meeting links are passed as URL parameters but only between local services

## License

This project is provided as-is for personal use. Modify and distribute according to your needs.

## Contributing

Feel free to submit issues, improvements, or feature requests. This is a personal project but contributions are welcome.

## Changelog

### Version 2.0
- Added comprehensive plugin settings configuration
- Implemented meeting link detection and automatic link creation
- Added customizable output format templates
- Added support for child blocks in templates
- Added configurable attendee name exclusion
- Added recurring event indicators
- Added bracket settings for creating page links
- Improved error handling and debug logging