from src.calendar.fetcher import get_outlook_events
from src.utils.text_cleaner import clean_body_text
import sys
import io
import logging
import os
import argparse
from datetime import datetime

# Create logs directory if it doesn't exist
os.makedirs('logs', exist_ok=True)

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

def parse_args():
    """Parse command-line arguments"""
    parser = argparse.ArgumentParser(description='Retrieve Outlook calendar events')
    parser.add_argument('--days-back', type=int, default=1, help='Number of days to look back')
    parser.add_argument('--days-forward', type=int, default=1, help='Number of days to look forward')
    return parser.parse_args()

def main():
    """Main entry point for the application"""
    # Parse command line arguments
    args = parse_args()
    
    # Get events from Outlook based on provided date range
    filtered_items = get_outlook_events(days_back=args.days_back, days_forward=args.days_forward)
    
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