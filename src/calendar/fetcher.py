import win32com.client
from datetime import datetime, timedelta
import logging

def get_outlook_events(days_back=1, days_forward=1):
    """Get Outlook calendar events within a date range"""
    logging.info(f"Retrieving Outlook calendar events for range: -{days_back} to +{days_forward} days")
    
    try:
        # Connect to Outlook
        logging.info("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)
        
        # Calculate date range
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        start_date = today - timedelta(days=days_back)
        end_date = today + timedelta(days=days_forward)
        logging.info(f"Date range set: {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}")
        
        # Format dates and create restriction
        start_str = start_date.strftime("%d/%m/%Y %H:%M %p")
        end_str = end_date.strftime("%d/%m/%Y %H:%M %p")
        
        # Get items
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        
        restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
        logging.info(f"Formatted restriction: {restriction}")
        
        filtered_items = items.Restrict(restriction)
        logging.info(f"Retrieved {filtered_items.Count if hasattr(filtered_items, 'Count') else '?'} items")
        
        return filtered_items
        
    except Exception as e:
        logging.error(f"Error accessing Outlook: {e}", exc_info=True)
        return []