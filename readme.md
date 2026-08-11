# Outlook Events Logseq Plugin (Database Version)

A Logseq plugin that fetches events from your local Outlook calendar and inserts them into your journal pages as native, typed database properties. This plugin consists of two components: a Logseq plugin (JavaScript) and a local API service (Python) that reads from your Outlook OST file.

> **Logseq Database version required.** This branch targets the Logseq **DB** graph. It uses the DB property APIs (`upsertProperty` / `upsertBlockProperty`) and will not work on the older markdown (file) version. For the markdown version, use the `main` branch.

## Features

- 📅 Fetches events from your local Outlook calendar
- 📝 Inserts events directly into Logseq journal pages
- ⏰ Displays events in chronological order (earliest to latest)
- 👥 Adds attendees as linked page nodes (with configurable exclusions)
- ⏱️ Calculates and displays event duration
- 🏷️ Writes structured data as native, typed Logseq DB properties
- 🔃 Shows recurring event indicator
- 🔗 Automatically detects and creates meeting links
- ⚙️ Configurable through Logseq plugin settings
- 🔀 Optional per-field toggles for location and description

## Components

### 1. Logseq Plugin (`main.js`)
The main plugin that runs within Logseq and provides the `/Get Events` slash command.

### 2. API Service (`outlook_events_api.py`)
A local Flask API service that connects to your Outlook installation and retrieves calendar events. Sends CORS headers so the Logseq DB plugin origin (`lsp://logseq.com`) can reach it.

### 3. Service Management Scripts
- `start_api_service.bat` - Starts the API service in background
- `launch_api.vbs` - VBScript for silent startup (used with Task Scheduler)

## Prerequisites

- **Logseq (Database version)** installed and running
- **Microsoft Outlook** installed on Windows with access to your calendar
- **Python 3.x** installed
- Required Python packages: `flask`, `pywin32`

## Installation

### Step 1: Set up Project Files

1. **Create the plugin directory:**
   - Open your Logseq graph/workspace folder (where your pages and journals are stored)
   - Create a new folder named `logseq-get-outlook-events` in the root of your Logseq directory
   - This keeps all plugin files organized together in your workspace

2. **Copy all project files to the plugin directory:**
   - `main.js` (Logseq plugin)
   - `outlook_events_api.py` (API service)
   - `requirements.txt` (Python dependencies)
   - `start_api_service.bat` (Service management script)
   - Any other related files

   Your folder structure should look like:
   ```
   YourLogseqWorkspace/
   ├── pages/
   ├── journals/
   ├── outlook-events-plugin/
   │   ├── main.js
   │   ├── outlook_events_api.py
   │   ├── requirements.txt
   │   ├── start_api_service.bat
   │   └── (other plugin files)
   └── (other Logseq files)
   ```

### Step 2: Set up the Python API Service

1. **Open the Command Prompt and navigate to the plugin directory:**
   ```bash
   cd YourLogseqWorkspace\logseq-get-outlook-events
   ```

2. **(Optional but Recommended) Create a virtual environment:**
   
   A virtual environment isolates this project's Python dependencies from your system-wide Python installation. This prevents version conflicts with other Python projects and keeps your system clean.
   
   ```bash
   # Create virtual environment in the plugin directory
   python -m venv .venv
   
   # Activate virtual environment
   .venv\Scripts\activate

   ```
   
   When the virtual environment is active, you'll see `(.venv)` in your command prompt.

3. **Install required Python packages:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Test the API service:**
   ```bash
   python outlook_events_api.py
   ```
   
   The service will start on `http://localhost:5000`. Press `Ctrl+C` to stop it for now.

**Note:** If you created a virtual environment, the `start_api_service.bat` script will automatically use it. For manual starts, remember to activate it (`.venv\Scripts\activate`) each time.

### Step 3: Install the Logseq Plugin (Database version)

The Logseq DB version removed the old **"Load unpacked plugin"** button from the main Plugins menu. Use one of the methods below.

> **Note on folder paths:** The DB plugin loader can fail to load a local plugin whose path contains spaces (e.g. a OneDrive folder). If your path has spaces, create a space-free junction that points to the plugin folder and use that path instead:
>
> ```powershell
> New-Item -ItemType Junction -Path "C:\logseq-plugins\get-outlook-events" -Target "C:\path\to\logseq-get-outlook-event"
> ```

#### Method A — plugins.edn (recommended)

1. Open (or create) the file `C:\Users\<you>\.logseq\config\plugins.edn`
2. Add an entry pointing to the plugin folder with a local `:url` (use the junction path if your real path has spaces):
   ```edn
   :logseq-get-outlook-events {:url "C:/logseq-plugins/get-outlook-events", :effect false, :theme false}
   ```
   > Do **not** use the marketplace `:repo`/`:version` format for a local plugin — that routes through `plugins.logseq.io` and fails with a GitHub/Cloudflare error.
3. Fully quit Logseq (system tray → Quit) and relaunch.

#### Method B — register from the developer console

If `plugins.edn` doesn't pick it up, open Logseq DevTools (`Ctrl+Shift+I`), make sure the console context is the main Logseq window (you should see `LSPluginCore` in `Object.keys(window)`), then run:

```javascript
LSPluginCore.register({ url: "C:/logseq-plugins/get-outlook-events" })
```

The plugin registers immediately and appears in the Plugins panel.

#### Verify it loaded

In the DevTools console you should see `Get Outlook Events Plugin Loaded`. If you instead see a message that the plugin **requires a Logseq DB build**, your Logseq version predates the DB property APIs — update Logseq.

### Step 4: (Optional) Set up Automatic API Service Startup

For convenience, you can configure the API service to start automatically when you log into Windows using Task Scheduler.

#### Configure Task Scheduler

1. **Open Task Scheduler:**
   - Press `Win + R`, type `taskschd.msc`, and press Enter
   - Or search for "Task Scheduler" in the Start menu

2. **Create a new task:**
   - In the right panel, click "Create Task..." (not "Create Basic Task")

3. **General Tab:**
   - Name: `Outlook Events API Service`
   - Description: `Automatically starts the Outlook Events API service for Logseq plugin`
   - Check "Run only when user is logged on"
   - Check "Run with highest privileges" (recommended for COM access)

4. **Triggers Tab:**
   - Click "New..."
   - Begin the task: "At log on"
   - Settings: "Specific user" (should show your username)
   - Check "Delay task for: 30 seconds" (gives system time to fully load)
   - Click "OK"

5. **Actions Tab:**
   - Click "New..."
   - Action: "Start a program"
   - Program/script: `cmd.exe`
   - Add arguments: `/c "C:\path\to\your\logseq\workspace\logseq-get-outlook-events\launch_api.vbs"` (replace with actual path)
   - Click "OK"

6. **Conditions Tab:**
   - Uncheck all boxes

7. **Settings Tab:**
   - Check "Allow task to be run on demand"
   - Check "If the task fails, restart every" and select 1 minute and 3 attempts in the dropdowns 
   - Check "If the running task does not end when requested, force it to stop"
   - Set "If task is already running, then the following rule applies" to "Do not start a new instance"
   - Click "OK"

8. **Save the Task**
   - Click "OK"

8. **Test the scheduled task:**
   - Right-click on your new task in the Task Scheduler Library
   - Select "Run" to test it
   - Check if the API service starts by visiting [`http://localhost:5000/health`](http://localhost:5000/health) in your browser.  You should get the following response:
   ```json
   { 
   "outlook_status": "running",
   "service": "Outlook Events API",
   "status": "healthy" 
   }
   ```
   - You can then check that the service is also returning events from your calendar by visiting [`http://localhost:5000/events/{date}`](http://localhost:5000/events/{date}), but replace "{date}" with a day you know you have an event on your calendar. The response should look something like this:
   ```json
   {
   "date": "2025-08-21",
   "events": [
      {
         "attendees": [
         "Name, Your",
         "Simpson, Homer"
         ],
         "description": "Agenda: \n\r Review current safety procedures for plant shutdown.",
         "end": "2025-08-21 15:30:00+00:00",
         "isRecurring": true,
         "location": "Microsoft Teams Meeting",
         "meetingLinks": [],
         "start": "2025-08-21 15:00:00+00:00",
         "subject": "Homer Touchbase"
      }
   ],
   "success": true
   }
   ```

#### Verify Automatic Startup

After setting up the scheduled task:

1. **Restart your computer** to test automatic startup
2. **Wait about 1-2 minutes** after logging in (allow time for the delayed start)
3. **Test the service** by opening a web browser and going to [`http://localhost:5000/health`](http://localhost:5000/health)
4. You should see:
```json
   { 
   "outlook_status": "running",
   "service": "Outlook Events API",
   "status": "healthy" 
   }
   ```

#### Managing the Automatic Service

- **To disable automatic startup:** Open Task Scheduler, find your task, right-click and select "Disable"
- **To modify the startup delay:** Edit the task, go to Triggers tab, and change the delay time
- **To stop the service manually:** Use the provided `stop_api_service.bat` file

## Usage

### Starting the Services

**If you set up automatic startup:** The API service should start automatically when you log in. Just open Logseq and start using the plugin.

**If you prefer manual startup:**

1. **Start the Python API service first:**
   ```bash
   # Option 1: Direct command
   python outlook_events_api.py
   
   # Option 2: Using the batch file (runs in background)
   start_api_service.bat
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
```

The block content holds the event title (optionally a page-linked node), the recurring emoji, and any detected meeting links. All structured data is written as **native, typed Logseq DB properties** on that block:

| Property | Type | Notes |
|----------|------|-------|
| `event-time` | text | Start time, formatted per your Time Format setting |
| `event-duration` | text | `HH:MM:SS` |
| `event-location` | text | Omitted when **Include Location** is off |
| `event-description` | text | Omitted when **Include Description** is off; truncated per Description Max Length |
| `recurring` | checkbox | Set for recurring events |
| `attendees` | node (many) | Each attendee becomes a linked page |

Where:
- 🔃 indicates a recurring event
- Meeting links are automatically detected and added as clickable links
- Attendees render as clickable page links (nodes), not plain text

> These properties are defined automatically the first time the plugin loads, via the DB `upsertProperty` API.

## Configuration

The plugin is fully configurable through Logseq's plugin settings. Access settings via:
**Settings → Plugins → Get Outlook Events**

### Available Settings

#### **Excluded User Name**
- **Description**: Enter your name as it is returned by Outlook to exclude it from attendees lists
- **Default**: Empty (no exclusion)
- **Example**: `"Simpson, Homer"` or `"Diana Prince"`

#### **API URL**
- **Description**: URL of the Outlook Events API service
- **Default**: `http://localhost:5000`
- **Usage**: Change if running the API on a different port or machine

#### **Add Double Brackets to Event Titles to create/add page references**
- **Setting name**: Event Title as Node
- **Options**: 
  - `all` - Make all event titles page-linked nodes
  - `recurring` - Only link recurring events
  - `none` - No linking (default)
- **Purpose**: Turns event titles into Logseq page nodes

> Attendees are **always** added as linked page nodes via the `attendees` property, so a separate attendee bracket setting is no longer needed.

#### **Time Format Configuration**

The plugin uses tokens to format dates and times. You can customize the format using the following tokens:

#### Date Tokens
- **Year**
  - `YYYY` - Full year (2025)
  - `YY` - Short year (25)
- **Month**
  - `MMMM` - Full month name (September)
  - `MMM` - Short month name (Sep)
  - `MM` - Two-digit month (09)
  - `M` - Month number (9)
- **Day**
  - `DD` - Two-digit day (15)
  - `D` - Day number (15)
  - `dddd` - Full day name (Monday)
  - `ddd` - Short day name (Mon)

#### Time Tokens
- **Hour**
  - `HH` - Two-digit 24-hour (09)
  - `H` - 24-hour (9)
  - `hh` - Two-digit 12-hour (09)
  - `h` - 12-hour (9)
- **Minute**
  - `mm` - Two-digit minutes (05)
  - `m` - Minutes (5)
- **Second**
  - `ss` - Two-digit seconds (03)
  - `s` - Seconds (3)
- **AM/PM**
  - `A` - Uppercase (AM)
  - `a` - Lowercase (am)

#### Examples
- `h:mm a` → 9:30 am
- `HH:mm` → 09:30
- `ddd, MMM D, h:mm a` → Mon, Sep 15, 9:30 AM
- `YYYY-MM-DD HH:mm` → 2025-09-15 09:30

#### **Meeting Base URLs**
- **Description**: Comma-separated list of base URLs to search for meeting links
- **Default**: `https://teams.microsoft.com,https://zoom.us,https://meet.google.com`
- **Example**: `https://teams.microsoft.com,https://zoom.us,https://webex.com`
- **Note**: The plugin will find any URLs in event descriptions that start with these base URLs

#### **Output Format**

The DB version writes fixed, native properties rather than a free-form text template. Instead of a template, use the toggles below to control which properties are written.

#### **Include Location**
- **Description**: Whether to write the `event-location` property
- **Default**: `true`

#### **Include Description**
- **Description**: Whether to write the `event-description` property
- **Default**: `true`

#### **Include Empty Fields**
- **Description**: Whether to write properties even when their value is empty
- **Default**: `false`
- **Example**: When `true`, writes `event-location` even if the event has no location

#### **Description Max Length**
- **Description**: Maximum characters to include from event descriptions (applies only when Include Description is on)
- **Default**: `200`
- **Usage**: Set to `0` for no limit

### Meeting Links Feature

The plugin automatically scans event descriptions for meeting links and adds clickable "Join Meeting" links to event titles. 

**How it works:**
1. Configure the meeting base URLs in settings (e.g., `https://teams.microsoft.com/l/meetup-join`)
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
http://localhost:5000/events/2025-08-19?meeting_urls=https://teams.microsoft.com/l/meetup-join
```

## Recurring Events

Recurring events are automatically identified and marked with a 🔃 emoji, and the `recurring` checkbox property is set. Use the **Event Title as Node** setting to page-link recurring events only.

## Troubleshooting

### Common Issues

1. **"No events found" message:**
   - Ensure the Python API service is running
   - Check that you're on a journal page (daily note)
   - Verify events exist in Outlook for that date
   - Occasionaly the API service disconnects from the Outlook COM.  Stopping and restarting the service should correct the issue

2. **API connection errors:**
   - Confirm the Python service is running on localhost:5000
   - Check the API URL setting in plugin configuration
   - Ensure Windows Firewall isn't blocking the connection

3. **Meeting links not appearing:**
   - Verify meeting base URLs are configured correctly
   - Check that meeting links in Outlook match the configured base URLs
   - Ensure the Python service is receiving the meeting URLs parameter

4. **Plugin not loading (DB version):**
   - Confirm you used a local `:url` in `plugins.edn` (not the `:repo`/`:version` marketplace format)
   - If your plugin path contains spaces, use a space-free junction path (see Installation Step 3)
   - Try registering directly from the console: `LSPluginCore.register({ url: "C:/logseq-plugins/get-outlook-events" })`
   - Review the developer console for JavaScript errors

5. **CORS error ("No 'Access-Control-Allow-Origin' header"):**
   - The API service must send CORS headers so the plugin origin `lsp://logseq.com` can reach it
   - Ensure you are running the current `outlook_events_api.py` (it adds these headers) and restart the service
   - Verify with: `(Invoke-WebRequest http://localhost:5000/health -UseBasicParsing).Headers['Access-Control-Allow-Origin']` should return `*`

6. **"requires a Logseq DB build" message:**
   - Your Logseq version predates the DB property APIs — update Logseq to a current Database build

7. **Automatic startup not working:**
   - Check Task Scheduler to ensure the task is enabled and configured correctly
   - Verify the path to `launch_api.vbs` is correct in the scheduled task
   - Try running the scheduled task manually to test it
   - Check Windows Event Viewer for any error messages related to the task

### Debug Information

The plugin logs detailed information to Logseq's console. Open Developer Tools (CTRL + Shift + i) to view debug output.

## Technical Details

### Time Handling

- Events are fetched in Outlook's local timezone
- Times are displayed in either 12-hour format (9:00 AM) or 24-hour format (09:00) based on plugin settings
- Duration is always calculated and displayed as HH:MM:SS format

### Data Flow

1. User triggers `/Get Events` command on a journal page
2. Plugin extracts the journal date and configuration settings
3. API request sent to Python service with meeting URL parameters
4. Python service queries Outlook COM interface and extracts meeting links
5. Events returned as JSON (with CORS headers) including meeting links
6. Plugin inserts a block per event and sets native DB properties via `upsertBlockProperty`

### Meeting Link Detection

The Python service uses regular expressions to find URLs in event descriptions that start with the configured base URLs. It handles:
- Multiple meeting links per event
- Various URL formats and parameters
- Cleaning up URLs (removing trailing punctuation)

### Service Management

The background service management system includes:
- **Process detection and PID tracking** for reliable service control
- **Automatic cleanup** of stale process files
- **Log rotation** to prevent log files from growing too large
- **Silent background operation** using VBScript for Task Scheduler integration

## Security Notes

- The API service only runs locally (localhost)
- No external network connections required
- Calendar data never leaves your local machine
- Uses Windows COM interface for secure Outlook access
- Meeting links are passed as URL parameters but only between local services

## License

This project is available under GNU General Public License version 3.

## Contributing

Feel free to submit issues, improvements, or feature requests. This is a personal project but contributions are welcome through a GitHub Issues submittion or Pull Request at [https://github.com/DavidJayMartin/logseq-get-outlook-event](https://github.com/DavidJayMartin/logseq-get-outlook-event).