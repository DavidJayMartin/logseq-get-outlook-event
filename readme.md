# Outlook Events Logseq Plugin

A Logseq plugin that fetches events from your local Outlook calendar and inserts them into your journal pages. This plugin consists of two components: a Logseq plugin (JavaScript) and a local API service (Python) that reads from your Outlook OST file.

## Features

- 📅 Fetches events from your local Outlook calendar
- 📝 Inserts events directly into Logseq journal pages
- ⏰ Displays events in chronological order (earliest to latest)
- 👥 Shows attendees for each event (excluding yourself)
- ⏱️ Calculates and displays event duration
- 🏷️ Uses Logseq properties for structured data

## Components

### 1. Logseq Plugin (`main.js`)
The main plugin that runs within Logseq and provides the `/Get Events` slash command.

### 2. API Service (`outlook_events_api.py`)
A local Flask API service that connects to your Outlook installation and retrieves calendar events.

## Prerequisites

- **Logseq** installed and running
- **Microsoft Outlook** installed on Windows with access to your calendar
- **Python 3.x** installed
- **Node.js** (for Logseq plugin development)

## Installation

### Step 1: Set up the Python API Service

1. Save the `outlook_events_api.py` and `requirements.txt` files to a directory of your choice

2. Install required Python packages:
   ```bash
   pip install -r requirements.txt
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

Each event is inserted as a block with the following format:

```
Event Title
event-time:: 9:00 AM
event-duration:: 01:00:00
attendees:: [John Doe], [Jane Smith]
```

### API Endpoints

The Python service provides several endpoints for testing:

- `GET /health` - Health check
- `GET /events/YYYY-MM-DD` - Get events for specific date
- `GET /events/today` - Get today's events
- `GET /events?date=YYYY-MM-DD` - Get events via query parameter

Example:
```
http://localhost:5000/events/2025-08-19
http://localhost:5000/events/today
```

## Configuration

### Excluding Attendees

The API service automatically excludes "Martin, David" from attendee lists. To change this, modify line 35 in `outlook_events_api.py`:

```python
# Change this line to exclude your name
if recipient.Name != "Your Name Here":
```

### API Port

The service runs on port 5000 by default. To change this, modify the last line in `outlook_events_api.py`:

```python
app.run(host='0.0.0.0', port=5000, debug=True)  # Change port here
```

If you change the port, also update the fetch URL in `main.js`:

```javascript
const response = await fetch(`http://localhost:5000/events/${dateString}`);
```

## Troubleshooting

### Common Issues

1. **"No events found" message:**
   - Ensure the Python API service is running
   - Check that you're on a journal page (daily note)
   - Verify events exist in Outlook for that date

2. **API connection errors:**
   - Confirm the Python service is running on localhost:5000
   - Check Windows Firewall settings
   - Ensure Outlook is installed and configured

3. **Plugin not loading:**
   - Verify the plugin folder structure
   - Check Logseq's plugin settings
   - Review browser console for JavaScript errors

### Debug Information

The plugin logs detailed information to the browser console. Open Developer Tools (F12) to view debug output.

## Technical Details

### Time Handling

- Events are fetched in Eastern Time (Outlook's local timezone)
- Times are displayed in 12-hour format (AM/PM)
- Duration is calculated as HH:MM:SS format

### Data Flow

1. User triggers `/Get Events` command on a journal page
2. Plugin extracts the journal date
3. API request sent to Python service
4. Python service queries Outlook COM interface
5. Events returned as JSON
6. Plugin formats and inserts events as Logseq blocks

## Security Notes

- The API service only runs locally (localhost)
- No external network connections required
- Calendar data never leaves your local machine
- Uses Windows COM interface for secure Outlook access

## License

This project is provided as-is for personal use. Modify and distribute according to your needs.

## Contributing

Feel free to submit issues, improvements, or feature requests. This is a personal project but contributions are welcome.