import sys
import os
import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime, timedelta
import pytz

# Add the project root to the path so we can import our modules
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))
from src.calendar.fetcher import OutlookCalendarFetcher
from src.utils.text_cleaner import clean_body_text

def calculate_duration(start, end):
    """Calculate duration of an event in hours"""
    if not start or not end:
        return 0
    
    duration = (end - start).total_seconds() / 3600
    return round(duration, 2)

def get_events_df(days_back=7):
    """Fetch events and convert to DataFrame with calculated fields"""
    try:
        import pythoncom
        pythoncom.CoInitialize()  # Initialize COM for this thread
        
        fetcher = OutlookCalendarFetcher()
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days_back)
        events = fetcher.fetch_events(start_date, end_date)
        
        # Convert to DataFrame with useful fields
        events_data = []
        for event in events:
            try:
                # Get start and end times
                start = event.start.astimezone(pytz.timezone('UTC'))
                end = event.end.astimezone(pytz.timezone('UTC'))
                
                # Clean the body content
                body_content = clean_body_text(event.body) if event.body else ""
                
                # Calculate duration
                duration = calculate_duration(start, end)
                
                # Safely extract organizer name
                organizer_name = "Unknown"
                try:
                    if hasattr(event.organizer, 'name'):
                        organizer_name = event.organizer.name
                    elif isinstance(event.organizer, str):
                        organizer_name = event.organizer
                except:
                    pass  # Keep default "Unknown"
                
                # More robust categories handling:
                
                # Get categories properly
                categories = ""
                if hasattr(event, 'categories'):
                    if isinstance(event.categories, list):
                        categories = ", ".join(event.categories)
                    else:
                        categories = event.categories
                
                events_data.append({
                    "subject": event.subject,
                    "start": start,
                    "end": end,
                    "duration": duration,
                    "organizer": organizer_name,
                    "categories": categories,
                    "is_recurring": event.is_recurring if hasattr(event, 'is_recurring') else False,
                    "day_of_week": start.strftime("%A"),
                    "body": body_content[:200] + "..." if len(body_content) > 200 else body_content
                })
            except Exception as e:
                print(f"Error processing event: {str(e)}")
                # Continue with next event rather than failing entire process
                continue
                
        return pd.DataFrame(events_data)
    
    finally:
        # Always uninitialize COM when we're done
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def main():
    st.set_page_config(
        page_title="Outlook Calendar Analysis",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("Outlook Calendar Time Analysis")
    
    # Sidebar filters
    st.sidebar.header("Filters")
    days_back = st.sidebar.slider("Days to analyze", min_value=1, max_value=30, value=7)
    
    # Load data first to get categories
    with st.spinner("Fetching calendar data..."):
        try:
            all_events_df = get_events_df(days_back)
            if all_events_df.empty:
                st.warning("No calendar events found in the selected period.")
                return
        except Exception as e:
            st.error(f"Error fetching calendar data: {str(e)}")
            return
    
    # Extract unique categories and prepare for multiselect
    # Split categories that might contain multiple values
    all_categories = set()
    for cats in all_events_df['categories'].dropna():
        if cats:  # Check if not empty
            # Split by comma and strip whitespace
            for cat in [c.strip() for c in cats.split(',')]:
                if cat:  # Only add non-empty categories
                    all_categories.add(cat)
    
    all_categories = sorted(list(all_categories))
    
    # Set default to all categories except OOO
    default_categories = [cat for cat in all_categories if cat != "OOO"]
    
    # Category filter in sidebar
    st.sidebar.header("Category Filter")
    selected_categories = st.sidebar.multiselect(
        "Select categories to include:",
        options=all_categories,
        default=default_categories
    )
    
    # Filter dataframe based on selected categories
    if selected_categories:
        # Filter to only include rows with at least one selected category
        mask = all_events_df['categories'].apply(
            lambda cats: any(cat in str(cats).split(',') for cat in selected_categories) if pd.notna(cats) else False
        )
        df = all_events_df[mask].copy()
    else:
        # If nothing selected, include all
        df = all_events_df.copy()
    
    # Show how many events were filtered
    st.sidebar.info(f"Showing {len(df)} of {len(all_events_df)} events")
    
    # Show total time metrics
    total_hours = df["duration"].sum()
    total_meetings = len(df)
    avg_duration = df["duration"].mean()
    
    # Rest of your visualization code remains the same
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Hours in Meetings", f"{total_hours:.1f}")
    col2.metric("Number of Meetings", total_meetings)
    col3.metric("Average Duration (hours)", f"{avg_duration:.2f}")
    
    # Time by day of week
    st.header("Time Spent by Day of Week")
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    day_df = df.groupby("day_of_week")["duration"].sum().reset_index()
    day_df["day_of_week"] = pd.Categorical(day_df["day_of_week"], categories=day_order, ordered=True)
    day_df = day_df.sort_values("day_of_week")
    
    fig1 = px.bar(
        day_df,
        x="day_of_week",
        y="duration",
        labels={"duration": "Hours", "day_of_week": "Day"},
        title="Meeting Hours by Day of Week"
    )
    st.plotly_chart(fig1)
    
    # Time by organizer
    st.header("Time by Meeting Organizer")
    organizer_df = df.groupby("organizer")["duration"].sum().reset_index().sort_values("duration", ascending=False).head(10)
    fig2 = px.pie(
        organizer_df,
        values="duration",
        names="organizer",
        title="Meeting Hours by Organizer (Top 10)"
    )
    st.plotly_chart(fig2)
    
    # Meeting timeline
    st.header("Meeting Timeline")
    timeline_df = df[["subject", "start", "end", "duration"]].sort_values("start")
    fig3 = px.timeline(
        timeline_df,
        x_start="start",
        x_end="end",
        y="subject",
        color="duration",
        labels={"subject": "Meeting", "duration": "Hours"},
        title="Meeting Timeline"
    )
    fig3.update_yaxes(autorange="reversed")
    st.plotly_chart(fig3)
    
    # Raw data table
    st.header("Meeting Details")
    st.dataframe(
        df[["subject", "start", "end", "duration", "organizer", "categories"]],
        hide_index=True
    )

if __name__ == "__main__":
    main()
