import win32com.client
from datetime import datetime, timedelta
import logging

def get_outlook_events(days_back=1, days_forward=1):
    """Get Outlook calendar events within a date range"""
    logging.info("Retrieving Outlook calendar events")
    
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)
        
        # Calculate date range
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        start_date = today - timedelta(days=days_back)
        end_date = today + timedelta(days=days_forward)
        
        # Format dates and create restriction
        start_str = start_date.strftime("%m/%d/%Y %H:%M %p")
        end_str = end_date.strftime("%m/%d/%Y %H:%M %p")
        
        # Get items
        items = calendar.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        
        restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
        filtered_items = items.Restrict(restriction)
        
        return filtered_items
        
    except Exception as e:
        logging.error(f"Error accessing Outlook: {e}", exc_info=True)
        return []