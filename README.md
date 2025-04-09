# Outlook Calendar Notes

A modern Python application to fetch, process, and analyze your Microsoft Outlook calendar events.

## 🌟 Features

- **Calendar Integration**: Retrieve events directly from your Outlook client
- **Smart Filtering**: Focus on relevant events with customizable date ranges
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
- [ ] Add more exporters (e.g., CSV, JSON)
- [ ] Integration with task management systems
- [ ] Calendar event search functionality

## 📜 License

MIT

## 🤝 Contributing

Contributions welcome! Please feel free to submit a Pull Request.