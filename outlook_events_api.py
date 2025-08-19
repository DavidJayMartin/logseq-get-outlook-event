# outlook_api.py
from flask import Flask, jsonify, request
from datetime import datetime
import win32com.client
import pythoncom
import traceback

app = Flask(__name__)

def get_events(date_str):
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
                # Get list of recipients/attendees
                attendees = []
                try:
                    for recipient in item.Recipients:
                        # Exclude "Martin, David" from the attendees list
                        if recipient.Name != "Martin, David":
                            attendees.append(recipient.Name)
                except:
                    # Some items might not have recipients
                    pass
                
                events.append({
                    "subject": item.Subject,
                    "start": str(item.Start),
                    "end": str(item.End),
                    "location": item.Location,
                    "attendees": attendees
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
    
    if not date_str:
        return jsonify({"success": False, "error": "Date parameter is required. Use ?date=YYYY-MM-DD"}), 400
    
    result = get_events(date_str)
    
    if result["success"]:
        return jsonify(result)
    else:
        return jsonify(result), 400

@app.route('/events/<date_str>', methods=['GET'])
def get_events_by_path(date_str):
    """Get events for a specific date via URL path"""
    result = get_events(date_str)
    
    if result["success"]:
        return jsonify(result)
    else:
        return jsonify(result), 400

@app.route('/events/today', methods=['GET'])
def get_today_events():
    """Get events for today"""
    today = datetime.now().strftime("%Y-%m-%d")
    result = get_events(today)
    
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
    print("  GET /events?date=YYYY-MM-DD - Get events by query parameter")
    print("  GET /events/YYYY-MM-DD - Get events by URL path")
    print("  GET /events/today - Get today's events")
    print()
    print("Example usage:")
    print("  http://localhost:5000/events?date=2025-08-18")
    print("  http://localhost:5000/events/2025-08-18")
    print("  http://localhost:5000/events/today")
    print()
    
    # Run the Flask app
    app.run(host='0.0.0.0', port=5000, debug=True)