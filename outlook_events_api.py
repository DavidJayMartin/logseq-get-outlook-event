# outlook_api.py
from flask import Flask, jsonify, request
from datetime import datetime
import win32com.client
import pythoncom
import traceback
import re
import subprocess
import time

app = Flask(__name__)

def is_outlook_running():
    """Check if Outlook is currently running"""
    try:
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq OUTLOOK.EXE'], 
                              capture_output=True, text=True)
        return 'OUTLOOK.EXE' in result.stdout
    except:
        return False

def start_outlook():
    """Attempt to start Outlook"""
    try:
        subprocess.Popen(['outlook.exe'])
        # Wait a moment for Outlook to start
        time.sleep(3)
        return True
    except:
        return False

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
        # Initialize COM for this thread with security settings
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        
        # Parse date
        date = datetime.strptime(date_str, "%Y-%m-%d")
        
        # Check if Outlook is running
        if not is_outlook_running():
            print("Outlook is not running. Attempting to start...")
            if not start_outlook():
                return {"success": False, "error": "Outlook is not running and could not be started automatically. Please start Outlook manually and try again."}
            
            # Wait a bit more for Outlook to fully initialize
            time.sleep(5)
        
        # Try different approaches to connect to Outlook
        outlook = None
        last_error = None
        
        # Method 1: Try direct dispatch first
        try:
            print("Attempting direct Outlook connection...")
            outlook = win32com.client.Dispatch("Outlook.Application")
            outlook = outlook.GetNamespace("MAPI")
            print("✓ Direct connection successful")
        except Exception as e:
            last_error = e
            print(f"✗ Direct connection failed: {e}")
        
        # Method 2: Try GetActiveObject (connects to existing instance)
        if outlook is None:
            try:
                print("Attempting to connect to existing Outlook instance...")
                outlook_app = win32com.client.GetActiveObject("Outlook.Application")
                outlook = outlook_app.GetNamespace("MAPI")
                print("✓ Connection to existing instance successful")
            except Exception as e:
                last_error = e
                print(f"✗ Connection to existing instance failed: {e}")
        
        # Method 3: Try with explicit CLSID
        if outlook is None:
            try:
                print("Attempting connection with explicit CLSID...")
                outlook_app = win32com.client.Dispatch("{0006F03A-0000-0000-C000-000000000046}")
                outlook = outlook_app.GetNamespace("MAPI")
                print("✓ CLSID connection successful")
            except Exception as e:
                last_error = e
                print(f"✗ CLSID connection failed: {e}")
        
        if outlook is None:
            # Provide detailed error information
            error_code = getattr(last_error, 'hresult', None) if last_error else None
            error_msg = str(last_error) if last_error else "Unknown error"
            
            if error_code == -2146959355:  # Server execution failed
                return {
                    "success": False, 
                    "error": "Could not connect to Outlook COM interface. This usually indicates a permissions issue.",
                    "solutions": [
                        "Try running the API service as Administrator (right-click → Run as administrator)",
                        "Close and restart Microsoft Outlook completely",
                        "Restart the API service after restarting Outlook",
                        "Check if Windows Defender or antivirus is blocking COM access"
                    ],
                    "technical_details": f"Error code: {error_code}, Message: {error_msg}"
                }
            elif error_code == -2147221021:  # Operation unavailable
                return {
                    "success": False, 
                    "error": "Outlook COM interface is currently unavailable.",
                    "solutions": [
                        "Restart Microsoft Outlook",
                        "Wait a few moments after starting Outlook before trying again",
                        "Try running both Outlook and the API service as Administrator"
                    ],
                    "technical_details": f"Error code: {error_code}, Message: {error_msg}"
                }
            else:
                return {
                    "success": False, 
                    "error": f"Failed to connect to Outlook after trying multiple methods.",
                    "solutions": [
                        "Run the API service as Administrator",
                        "Restart both Outlook and the API service",
                        "Check Windows Event Log for COM errors",
                        "Verify Outlook is not in safe mode"
                    ],
                    "technical_details": f"Error code: {error_code}, Message: {error_msg}"
                }
        
        try:
            calendar = outlook.GetDefaultFolder(9)  # olFolderCalendar
        except Exception as e:
            return {
                "success": False,
                "error": "Connected to Outlook but could not access calendar folder.",
                "solutions": [
                    "Ensure Outlook has finished loading completely",
                    "Check that you have a valid Exchange/IMAP account configured",
                    "Try restarting Outlook and waiting for it to fully sync"
                ],
                "technical_details": str(e)
            }

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
    outlook_status = "running" if is_outlook_running() else "not running"
    return jsonify({
        "status": "healthy", 
        "service": "Outlook Events API",
        "outlook_status": outlook_status
    })

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
    print("  GET /health - Health check (includes Outlook status)")
    print("  GET /events?date=YYYY-MM-DD&meeting_urls=url1,url2 - Get events by query parameter")
    print("  GET /events/YYYY-MM-DD?meeting_urls=url1,url2 - Get events by URL path")
    print("  GET /events/today?meeting_urls=url1,url2 - Get today's events")
    print()
    print("Example usage:")
    print("  http://localhost:5000/events?date=2025-08-18&meeting_urls=https://teams.microsoft.com,https://zoom.us")
    print("  http://localhost:5000/events/2025-08-18?meeting_urls=https://meet.google.com")
    print("  http://localhost:5000/events/today")
    print()
    print("Checking Outlook status...")
    if is_outlook_running():
        print("✓ Outlook is currently running")
    else:
        print("⚠ Outlook is not running - it will be started automatically when needed")
    print()
    
    # Run the Flask app
    app.run(host='0.0.0.0', port=5000, debug=True)