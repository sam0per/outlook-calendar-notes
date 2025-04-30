from src.calendar.fetcher import OutlookCalendarFetcher
from src.utils.text_cleaner import clean_body_text
import sys
import io
import logging
import os
import argparse
import time
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
    parser.add_argument('--export-json', action='store_true', help='Export events to JSON file')
    parser.add_argument('--export-dir', type=str, default='exports', help='Directory to save exported JSON files')
    return parser.parse_args()

def sync_outlook():
    """Force Outlook to synchronize before fetching events"""
    try:
        import win32com.client
        logging.info("Initializing Outlook and forcing synchronization...")
        
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Access calendar folder to trigger sync
        calendar = namespace.GetDefaultFolder(9)  # 9 is the calendar folder
        
        # Force sync by accessing items
        _ = calendar.Items.Count
        
        # Add a small delay to allow sync to complete
        time.sleep(2)
        
        logging.info("Outlook synchronization completed")
        return True
    except Exception as e:
        logging.error(f"Error synchronizing Outlook: {str(e)}")
        return False

def main():
    """Main entry point for the application"""
    # Parse command line arguments
    args = parse_args()
    
    # Synchronize Outlook first
    sync_outlook()
    
    # Initialize the fetcher and get events
    fetcher = OutlookCalendarFetcher()
    filtered_items = fetcher.get_outlook_events(days_back=args.days_back, days_forward=args.days_forward)
    
    # Process events
    events = []
    for item in filtered_items:
        try:
            # Skip events with OOO category
            if hasattr(item, 'Categories') and item.Categories and "OOO" in item.Categories:
                logging.info(f"Skipping OOO event: {item.Subject}")
                continue
            
            # Validate required fields
            if not hasattr(item, 'Start') or not hasattr(item, 'End') or not item.Start or not item.End:
                logging.warning(f"Skipping event with missing dates: {item.Subject if hasattr(item, 'Subject') else 'Unknown'}")
                continue
                
            logging.info(f"Found event: {item.Subject}")
            
            event = {
                "Subject": item.Subject if hasattr(item, 'Subject') else "Untitled Event",
                # Convert COM datetime objects to standard Python datetime objects
                "Start": datetime(
                    item.Start.year, item.Start.month, item.Start.day, 
                    item.Start.hour, item.Start.minute, item.Start.second
                ) if hasattr(item, 'Start') and item.Start else None,
                "End": datetime(
                    item.End.year, item.End.month, item.End.day, 
                    item.End.hour, item.End.minute, item.End.second
                ) if hasattr(item, 'End') and item.End else None,
                "Location": item.Location if hasattr(item, 'Location') else "",
                "Body": clean_body_text(item.Body) if hasattr(item, 'Body') else "",
                "Categories": item.Categories if hasattr(item, 'Categories') else ""
            }
            events.append(event)
        except Exception as e:
            logging.error(f"Error processing event: {str(e)}")
            continue
    
    logging.info(f"Retrieved {len(events)} events")
    
    # Run exporter if requested
    if args.export_json and events:
        from src.exporters.json_exporter import JsonExporter
        import pandas as pd
        
        # Convert events to DataFrame
        events_df = pd.DataFrame(events)
        
        # Export events to JSON
        exporter = JsonExporter(output_dir=args.export_dir)
        try:
            filename = exporter.export_events(events_df, filename=f"calendar_events_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
            logging.info(f"Events exported to {filename}")
        except Exception as e:
            logging.error(f"Failed to export events: {str(e)}")
        
    # Just display summary instead of all event details
    if events:
        print(f"\nFound {len(events)} calendar events.")
    else:
        print("No events found or error occurred")

if __name__ == "__main__":
    main()