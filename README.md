# AI-Powered Calendar & Report Assistant

A modern Python application to fetch, process, and analyze your Microsoft Outlook calendar events.

## 🌟 Features

- **Calendar Integration**: Retrieve events directly from your Outlook client
- **Smart Filtering**: Focus on relevant events with customizable date ranges
- **Enhanced Synchronization**: Ensure all calendar events are properly collected with retry logic
- **Calendar Selection**: Specify which Outlook calendar to pull events from
- **Content Cleanup**: Automatically remove Teams boilerplate and format event details
- **Multiple Export Formats**: Console, Markdown, HTML (extensible)
- **UTF-8 Support**: Properly handle international characters and emojis
- **Comprehensive Logging**: Track operations with detailed logs

## 📋 Requirements

- Windows with Microsoft Outlook installed
- Python 3.6+
- Required packages (see `requirements.txt`)

## 🚀 Installation

1. Clone this repository:  
```bash
git clone https://github.com/yourusername/outlook-calendar-notes.git
cd outlook-calendar-notes
```

2. Install required packages:  
```bash
pip install -r requirements.txt
```

## 📊 Project Structure

```bash
outlook-calendar-notes/
├── src/                      # Source code
│   ├── calendar/             # Calendar interaction
│   ├── utils/                # Utility functions
│   ├── exporters/            # Output formatters
│   └── streamlit/            # Streamlit app (if applicable)
├── requirements.txt          # Python dependencies
├── tests/                    # Unit tests
├── data/                     # Saved outputs/cache
├── logs/                     # Log files
├── config/                   # Configuration
├── main.py                   # Entry point
└── README.md                 # Documentation
```

## 💻 Usage

Run the main script:

```bash
python main.py
```

For specific date ranges:

```bash
python main.py --days-back 3 --days-forward 7
```

This will fetch events from the last 3 days to the next 7 days.  
You can also export the results to a JSON file by using the `--export-json` flag.
    
```bash
python main.py --days-back 6 --days-forward 2 --export-json
```

The exported file will be saved in the `exports/` directory with a timestamp.

### Synchronization Options

To ensure all calendar events are properly collected, use these options:

```bash
python main.py --sync-timeout 15 --sync-retries 5 --force-full-sync
```

- `--sync-timeout`: Time in seconds to wait for synchronization (default: 10)
- `--sync-retries`: Number of times to retry synchronization if it fails (default: 3)
- `--force-full-sync`: Attempt a more thorough synchronization of your Outlook calendar

### Calendar Selection

To specify a different calendar folder (other than your default):

```bash
python main.py --calendar-name "Work Calendar"
```

This allows you to pull events from specific calendar folders in your Outlook.

## 🔧 Development

The modular architecture makes it easy to extend:

1. Add new exporters in `src/exporters/`
2. Create custom filters in `src/calendar/`
3. Enhance text processing in `src/utils/`

## 📝 Logging

Logs are stored in the `logs/` directory with comprehensive information about each operation.

## 🔜 Roadmap

- [x] Web interface with ~Flask~ Streamlit
- [x] Meeting analytics and statistics
- [x] Add exporter to JSON format
- [x] Enhanced calendar synchronization with retry logic
- [x] Support for multiple calendar folders
- [ ] Integration with task management systems
- [ ] Calendar event search functionality

## 📜 License

MIT

## 🤝 Contributing

Contributions welcome! Please feel free to submit a Pull Request.
