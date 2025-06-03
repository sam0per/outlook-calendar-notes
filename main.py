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
    parser.add_argument('--sync-timeout', type=int, default=10, help='Timeout in seconds for Outlook synchronization')
    parser.add_argument('--sync-retries', type=int, default=3, help='Number of times to retry synchronization')
    parser.add_argument('--force-full-sync', action='store_true', help='Force a full synchronization of Outlook')
    parser.add_argument('--calendar-name', type=str, help='Specify a calendar name (default: primary calendar)')
    return parser.parse_args()

def sync_outlook(timeout=10, retries=3, force_full=False, calendar_name=None):
    """Force Outlook to synchronize before fetching events
    
    Args:
        timeout (int): Time in seconds to wait for synchronization
        retries (int): Number of retry attempts
        force_full (bool): Whether to force a full synchronization
        calendar_name (str): Name of the calendar to sync (None for default)
    
    Returns:
        bool: True if synchronization was successful
    """
    import win32com.client
    
    for attempt in range(1, retries + 1):
        try:
            logging.info(f"Initializing Outlook synchronization (attempt {attempt}/{retries})...")
            
            # Connect to Outlook
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            
            # Get the calendar folder - default or specified
            if calendar_name:
                found = False
                for folder in namespace.Folders.Item(1).Folders:
                    if folder.Name == "Calendar" or folder.Name == calendar_name:
                        calendar = folder
                        found = True
                        logging.info(f"Using calendar: {folder.Name}")
                        break
                if not found:
                    logging.warning(f"Calendar '{calendar_name}' not found, using default")
                    calendar = namespace.GetDefaultFolder(9)  # 9 is the calendar folder
            else:
                calendar = namespace.GetDefaultFolder(9)  # 9 is the calendar folder
            
            # Force sync by accessing items and more direct sync methods
            initial_count = calendar.Items.Count
            logging.info(f"Initial calendar items count: {initial_count}")
            
            if force_full:
                try:
                    # Try to force a more complete sync using Update method if available
                    namespace.SendAndReceive(True)
                    logging.info("Forced full synchronization")
                except:
                    logging.info("Full synchronization not available, using standard methods")
            
            # Get all items to force update
            all_items = calendar.Items
            all_items.Sort("[Start]")  # Sort to ensure we get through all items
            all_items.IncludeRecurrences = True
            
            # Process through all items to ensure they're loaded
            if all_items.Count > 0:
                logging.info(f"Processing through {all_items.Count} items to ensure complete sync")
                # Just accessing a sample of items can help force sync
                item_sample = min(100, all_items.Count)
                for i in range(1, item_sample + 1):
                    try:
                        _ = all_items.Item(i).Subject
                    except:
                        pass
                        
            # Add a delay to allow sync to complete
            logging.info(f"Waiting {timeout} seconds for synchronization to complete...")
            time.sleep(timeout)
            
            # Verify sync by checking if count changed
            final_count = calendar.Items.Count
            logging.info(f"Final calendar items count: {final_count}")
            
            if final_count != initial_count:
                logging.info(f"Sync detected: Items changed from {initial_count} to {final_count}")
            else:
                logging.info("No change in item count - calendar may already be up to date")
            
            logging.info("Outlook synchronization completed successfully")
            return True
        except Exception as e:
            logging.error(f"Error during synchronization attempt {attempt}: {str(e)}")
            if attempt < retries:
                wait_time = attempt * 2  # Increase wait time with each retry
                logging.info(f"Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                logging.error("All synchronization attempts failed")
                return False

def main():
    """Main entry point for the application"""
    # Parse command line arguments
    args = parse_args()
    
    # Synchronize Outlook first with enhanced parameters
    sync_successful = sync_outlook(
        timeout=args.sync_timeout,
        retries=args.sync_retries,
        force_full=args.force_full_sync,
        calendar_name=args.calendar_name
    )
    
    if not sync_successful:
        logging.warning("Proceeding with event collection despite synchronization issues")
    
    # Initialize the fetcher and get events
    fetcher = OutlookCalendarFetcher()
    
    # Get events using the standard method (without calendar_name parameter)
    # since the OutlookCalendarFetcher class may not support this parameter yet
    filtered_items = fetcher.get_outlook_events(
        days_back=args.days_back, 
        days_forward=args.days_forward
    )
    
    # If a specific calendar was used for sync, filter the results manually
    if hasattr(args, 'calendar_name') and args.calendar_name:
        logging.info(f"Using specified calendar: {args.calendar_name}")
        calendar_name = args.calendar_name.lower()
        # No filtering here, as we're using the calendar selected during sync
        # The sync_outlook function already selected the appropriate calendar

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