import win32com.client
from datetime import datetime, timedelta
import logging
import sys
import io
import re

# Fix console encoding for Unicode support
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Set up logging with UTF-8 encoding
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('outlook_calendar.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def clean_body_text(body):
    """Remove Microsoft Teams help information and other unnecessary content"""
    if not body:
        return ""
    
    # Remove everything after "Need help?" line
    body = re.split(r"Need help\?.*?<https://aka\.ms/JoinTeamsMeeting\?omkt=.*?>", body, flags=re.DOTALL)[0]
    
    # Additional cleanup: remove common meeting footers
    patterns = [
        r"Microsoft Teams.*?(?:\r\n|\n).*?Join conversation",
        r"________________+.*$",  # Common email/calendar separators
        r"Click here to join.*$",
        r"Join with a video conferencing.*$",
        r"Join Microsoft Teams Meeting.*$",
    ]
    
    for pattern in patterns:
        body = re.split(pattern, body, flags=re.DOTALL)[0]
    
    # Trim whitespace and remove extra blank lines
    body = re.sub(r'\n{3,}', '\n\n', body.strip())
    
    return body

def get_today_events():
    """Get only events scheduled for today from Outlook calendar"""
    logging.info("Starting to retrieve today's Outlook calendar events")
    
    try:
        # Connect to Outlook
        logging.info("Connecting to Outlook")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Access the calendar folder
        logging.info("Accessing calendar folder")
        calendar = namespace.GetDefaultFolder(9)  # 9 is the enum for calendar folder
        
        # Calculate today's date range
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        tomorrow = today + timedelta(days=1)
        yesterday = today - timedelta(days=1)
        logging.info(f"Today's date: {today}")
        
        # Format dates for restriction
        start_str = yesterday.strftime("%d/%m/%Y %H:%M %p")
        end_str = tomorrow.strftime("%d/%m/%Y %H:%M %p")
        logging.info(f"Date range: {start_str} to {end_str}")
        
        # Get calendar items
        items = calendar.Items
        items.Sort("[Start]")  # Sort by start time
        items.IncludeRecurrences = True
        
        # Create a restriction to filter by date range
        restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
        logging.info(f"Using restriction: {restriction}")
        filtered_items = items.Restrict(restriction)
        
        # Process events
        events = []
        logging.info("Processing calendar items")
        for item in filtered_items:
            # Skip events with OOO category
            if item.Categories and "OOO" in item.Categories:
                logging.info(f"Skipping OOO event: {item.Subject}")
                continue
            
            logging.info(f"Found event: {item.Subject}")
            event = {
            "Subject": item.Subject,
            "Start": item.Start,
            "End": item.End,
            "Location": item.Location,
            "Body": clean_body_text(item.Body),
            "Categories": item.Categories
            }
            events.append(event)
        
        logging.info(f"Retrieved {len(events)} events for today")
        return events
        
    except Exception as e:
        logging.error(f"Error accessing Outlook: {e}", exc_info=True)
        return []

if __name__ == "__main__":
    events = get_today_events()
    
    if events:
        print(f"\nFound {len(events)} calendar events for today:")
        for i, event in enumerate(events, 1):
            print(f"\n--- Event {i} ---")
            print(f"Subject: {event['Subject']}")
            print(f"Start: {event['Start']}")
            print(f"End: {event['End']}")
            print(f"Location: {event['Location']}")
            print(f"Body: {event['Body']}")
            print(f"Categories: {event['Categories']}")
    else:
        print("No events found for today or error occurred")