import win32com.client
import pythoncom  # Add this import
from datetime import datetime, timedelta
import logging

class OutlookCalendarFetcher:
    """Class to fetch calendar events from Outlook"""
    
    def __init__(self):
        """Initialize the Outlook connection"""
        self.outlook = None
        self.namespace = None
        self.calendar = None
        
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.calendar = self.namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
        except Exception as e:
            logging.error(f"Failed to initialize Outlook connection: {e}", exc_info=True)
    
    def fetch_events(self, start_date, end_date):
        """Fetch events between the specified dates"""
        if not all([self.outlook, self.namespace, self.calendar]):
            # Try to initialize again in case this is called from a different thread
            try:
                pythoncom.CoInitialize()
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                self.namespace = self.outlook.GetNamespace("MAPI")
                self.calendar = self.namespace.GetDefaultFolder(9)
            except Exception as e:
                logging.error(f"Failed to reinitialize Outlook connection: {e}", exc_info=True)
                return []
            
        try:
            # Format dates and create restriction
            start_str = start_date.strftime("%d/%m/%Y %H:%M %p")
            end_str = end_date.strftime("%d/%m/%Y %H:%M %p")
            
            # Get items
            items = self.calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True
            
            restriction = f"[Start] >= '{start_str}' AND [Start] < '{end_str}'"
            logging.info(f"Formatted restriction: {restriction}")
            
            filtered_items = items.Restrict(restriction)
            logging.info(f"Retrieved {filtered_items.Count if hasattr(filtered_items, 'Count') else '?'} items")
            
            # Convert to a list to make it easier to work with
            events = []
            for item in filtered_items:
                events.append(item)
            
            return events
            
        except Exception as e:
            logging.error(f"Error fetching Outlook events: {e}", exc_info=True)
            return []
        
    def __del__(self):
        """Cleanup COM resources when the object is destroyed"""
        try:
            pythoncom.CoUninitialize()
        except:
            pass
            
    def get_outlook_events(self, days_back=1, days_forward=1):
        """Legacy method for backwards compatibility"""
        today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        start_date = today - timedelta(days=days_back)
        end_date = today + timedelta(days=days_forward)
        
        return self.fetch_events(start_date, end_date)