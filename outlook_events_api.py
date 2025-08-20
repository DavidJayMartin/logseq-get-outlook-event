# outlook_api.py
from flask import Flask, jsonify, request
from datetime import datetime
import win32com.client
import pythoncom
import traceback
import re

app = Flask(__name__)

def extract_meeting_links(body_text, base_urls=None):
    """Extract meeting links from event body text"""
    if not body_text or not base_urls:
        return []
    
    meeting_links = []
    
    # Convert base URLs to a list if it's a string
    if isinstance(base_urls, str):
        base_urls = [url.strip() for url in base_urls.split(',') if url.strip()]
    
    for base_url in base_urls:
        base_url = base_url.strip()
        if not base_url:
            continue
            
        # Create regex pattern to find URLs starting with the base URL
        # This will match the base URL followed by any valid URL characters
        pattern = re.escape(base_url) + r'[^\s<>\"\'\]\)\}]*'
        
        # Find all matches in the body text
        matches = re.findall(pattern, body_text, re.IGNORECASE)
        
        for match in matches:
            # Clean up the URL (remove trailing punctuation that might not be part of the URL)
            cleaned_url = re.sub(r'[.,;!?\]\)\}]+$', '', match)
            if cleaned_url and cleaned_url not in meeting_links:
                meeting_links.append(cleaned_url)
    
    return meeting_links

def get_events(date_str, meeting_base_urls=None):
    """Get Outlook events for a specific date"""
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
        # Parse date
        date = datetime.strptime(date_str, "%Y-%m-%d")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)  # olFolderCalendar

        # Filter appointments - use a broader range first
        begin = date.strftime("%m/%d/%Y 00:00 AM")
        end = date.strftime("%m/%d/%Y 11:59 PM")
        items = calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        restriction = f"[Start] >= '{begin}' AND [Start] <= '{end}'"
        restricted_items = items.Restrict(restriction)

        events = []
        for item in restricted_items:
            # Additional filtering to ensure we only get events for the specified date
            item_start = item.Start
            item_date = datetime(item_start.year, item_start.month, item_start.day)
            
            # Only include events that actually start on the requested date
            if item_date == date:
                # Get list of recipients/attendees (no filtering - let JavaScript handle it)
                attendees = []
                try:
                    for recipient in item.Recipients:
                        attendees.append(recipient.Name)
                except:
                    # Some items might not have recipients
                    pass
                
                # Extract event body and meeting links
                event_body = item.Body if hasattr(item, 'Body') else ""
                meeting_links = extract_meeting_links(event_body, meeting_base_urls)
                
                events.append({
                    "subject": item.Subject,
                    "start": str(item.Start),
                    "end": str(item.End),
                    "location": item.Location,
                    "attendees": attendees,
                    "isRecurring": item.IsRecurring if hasattr(item, 'IsRecurring') else False,
                    "description": event_body,
                    "meetingLinks": meeting_links
                })

        return {"success": True, "events": events, "date": date_str}
    
    except ValueError as e:
        return {"success": False, "error": f"Invalid date format. Use YYYY-MM-DD. Error: {str(e)}"}
    except Exception as e:
        return {"success": False, "error": f"Failed to retrieve events: {str(e)}", "traceback": traceback.format_exc()}
    finally:
        # Always uninitialize COM when done
        try:
            pythoncom.CoUninitialize()
        except:
            pass

@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({"status": "healthy", "service": "Outlook Events API"})

@app.route('/events', methods=['GET'])
def get_events_api():
    """Get events for a specific date via query parameter"""
    date_str = request.args.get('date')
    meeting_base_urls = request.args.get('meeting_urls', '')
    
    if not date_str:
        return jsonify({"success": False, "error": "Date parameter is required. Use ?date=YYYY-MM-DD"}), 400
    
    result = get_events(date_str, meeting_base_urls)
    
    if result["success"]:
        return jsonify(result)
    else:
        return jsonify(result), 400

@app.route('/events/<date_str>', methods=['GET'])
def get_events_by_path(date_str):
    """Get events for a specific date via URL path"""
    meeting_base_urls = request.args.get('meeting_urls', '')
    result = get_events(date_str, meeting_base_urls)
    
    if result["success"]:
        return jsonify(result)
    else:
        return jsonify(result), 400

@app.route('/events/today', methods=['GET'])
def get_today_events():
    """Get events for today"""
    today = datetime.now().strftime("%Y-%m-%d")
    meeting_base_urls = request.args.get('meeting_urls', '')
    result = get_events(today, meeting_base_urls)
    
    if result["success"]:
        return jsonify(result)
    else:
        return jsonify(result), 400

@app.errorhandler(404)
def not_found(error):
    return jsonify({"success": False, "error": "Endpoint not found"}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({"success": False, "error": "Internal server error"}), 500

if __name__ == '__main__':
    print("Starting Outlook Events API...")
    print("Available endpoints:")
    print("  GET /health - Health check")
    print("  GET /events?date=YYYY-MM-DD&meeting_urls=url1,url2 - Get events by query parameter")
    print("  GET /events/YYYY-MM-DD?meeting_urls=url1,url2 - Get events by URL path")
    print("  GET /events/today?meeting_urls=url1,url2 - Get today's events")
    print()
    print("Example usage:")
    print("  http://localhost:5000/events?date=2025-08-18&meeting_urls=https://teams.microsoft.com,https://zoom.us")
    print("  http://localhost:5000/events/2025-08-18?meeting_urls=https://meet.google.com")
    print("  http://localhost:5000/events/today")
    print()
    
    # Run the Flask app
    app.run(host='0.0.0.0', port=5000, debug=True)