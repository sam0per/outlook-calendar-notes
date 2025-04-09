from src.calendar.fetcher import get_outlook_events
from src.utils.text_cleaner import clean_body_text
import sys
import io
import logging
from datetime import datetime

# Fix console encoding for Unicode support
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Set up logging with UTF-8 encoding
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/outlook_calendar.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def main():
    """Main entry point for the application"""
    # Get today's events from Outlook
    filtered_items = get_outlook_events(days_back=1, days_forward=1)
    
    # Process events
    events = []
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
    
    logging.info(f"Retrieved {len(events)} events")
    
    # Display events
    if events:
        print(f"\nFound {len(events)} calendar events:")
        for i, event in enumerate(events, 1):
            print(f"\n--- Event {i} ---")
            print(f"Subject: {event['Subject']}")
            print(f"Start: {event['Start']}")
            print(f"End: {event['End']}")
            print(f"Location: {event['Location']}")
            print(f"Body: {event['Body']}")
            print(f"Categories: {event['Categories']}")
    else:
        print("No events found or error occurred")

if __name__ == "__main__":
    main()