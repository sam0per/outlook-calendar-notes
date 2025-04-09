import json
import os
import logging
from datetime import datetime
from typing import Optional, Dict, Any, List
import pandas as pd


class JsonExporter:
    """Exports calendar events to JSON format suitable for LLM input"""

    DEFAULT_PROMPT_TEMPLATE = """
    # Calendar Event Analysis
    Below is a JSON representation of my calendar events. Please analyze this data and provide insights about:
    1. How I'm spending my time
    2. Patterns in my meetings and appointments
    3. Suggestions for improving my calendar management

    Calendar Events:
    ```json
    {events}
    ```
    What insights can you provide based on this calendar data?"""

    def __init__(self, output_dir: str = "exports", log_level: str = "INFO"):
        """
        Initialize the exporter with output directory and logging configuration

        Args:
            output_dir: Directory to save exported files
            log_level: Logging level ('DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL')
        """
        self.output_dir = output_dir
        self._setup_logging(log_level)
        self.logger = logging.getLogger(__name__)

        try:
            os.makedirs(output_dir, exist_ok=True)
            self.logger.info(f"Output directory initialized at: {output_dir}")
        except OSError as e:
            self.logger.error(
                f"Failed to create output directory '{output_dir}': {str(e)}"
            )
            raise RuntimeError(f"Could not initialize output directory: {str(e)}")

    def _setup_logging(self, log_level: str):
        """Configure logging settings"""
        logging.basicConfig(
            level=log_level,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            handlers=[
                logging.StreamHandler(),
                logging.FileHandler("calendar_exporter.log"),
            ],
        )

    def _convert_datetime_fields(self, row: Dict[str, Any]) -> Dict[str, Any]:
        """Convert datetime fields in a row to ISO format strings"""
        converted = {}
        for col, val in row.items():
            try:
                if val is None:
                    converted[col] = None
                elif isinstance(val, (datetime, pd.Timestamp)):
                    # Handle timezone-aware datetime objects safely
                    converted[col] = val.isoformat()
                elif hasattr(val, 'timetuple'):  # Handle other datetime-like objects
                    # Convert to standard datetime then to ISO format
                    dt_val = datetime(*val.timetuple()[:6])
                    converted[col] = dt_val.isoformat()
                elif isinstance(val, list):
                    converted[col] = [
                        v.isoformat() if isinstance(v, (datetime, pd.Timestamp)) else
                        datetime(*v.timetuple()[:6]).isoformat() if hasattr(v, 'timetuple') else v
                        for v in val
                    ]
                else:
                    converted[col] = val
            except Exception as e:
                # Log the specific error with field details
                self.logger.warning(f"Failed to convert field '{col}' (type: {type(val)}): {str(e)}")
                converted[col] = str(val) if val is not None else None  # Convert to string as fallback
        return converted

    def export_events(
        self, events_df: pd.DataFrame, filename: Optional[str] = None
    ) -> str:
        """
        Export events DataFrame to a JSON file with progress tracking

        Args:
            events_df: Pandas DataFrame with event data
            filename: Optional custom filename

        Returns:
            Path to the exported file

        Raises:
            ValueError: If input data is invalid
            RuntimeError: If export fails
        """
        self.logger.info("Starting events export process")

        # Input validation
        if not isinstance(events_df, pd.DataFrame):
            error_msg = "Input must be a pandas DataFrame"
            self.logger.error(error_msg)
            raise ValueError(error_msg)

        if events_df.empty:
            self.logger.warning("Empty DataFrame provided for export")

        # Prepare filename
        filename = (
            filename
            or f"calendar_events_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        )
        output_path = os.path.join(self.output_dir, filename)
        self.logger.info(f"Preparing to export to: {output_path}")

        # Convert data with progress tracking
        events_list = []
        total_events = len(events_df)

        self.logger.info(f"Processing {total_events} events...")
        for i, (_, row) in enumerate(events_df.iterrows(), 1):
            try:
                event_dict = self._convert_datetime_fields(row.to_dict())
                events_list.append(event_dict)

                # Log progress every 10% or for every event if <10
                log_interval = max(1, total_events // 10)
                if i % log_interval == 0 or i == total_events:
                    self.logger.info(
                        f"Processed {i}/{total_events} events ({i/total_events:.0%})"
                    )
            except Exception as e:
                self.logger.error(f"Failed to process row {i}: {str(e)}")
                continue  # Skip bad rows but continue processing

        # Prepare export structure
        export_data = {
            "metadata": {
                "exported_at": datetime.now().isoformat(),
                "event_count": len(events_list),
                "original_count": total_events,
                "success_rate": (
                    f"{len(events_list)/total_events:.1%}"
                    if total_events > 0
                    else "N/A"
                ),
                "description": "Calendar events exported from Outlook",
            },
            "events": events_list,
        }

        # Write to file
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
            self.logger.info(
                f"Successfully exported {len(events_list)} events to {output_path}"
            )
            return output_path
        except (IOError, OSError, TypeError) as e:
            error_msg = f"Failed to write JSON file: {str(e)}"
            self.logger.error(error_msg)
            raise RuntimeError(error_msg)
        except Exception as e:
            error_msg = f"Unexpected error during export: {str(e)}"
            self.logger.critical(error_msg)
            raise RuntimeError(error_msg)

    def generate_llm_prompt(
        self, events_df: pd.DataFrame, prompt_template: Optional[str] = None
    ) -> str:
        """
        Generate a prompt for language models with event data

        Args:
            events_df: Pandas DataFrame with event data
            prompt_template: Optional template string with {events} placeholder

        Returns:
            String with formatted prompt including event data

        Raises:
            ValueError: If input data is invalid
            RuntimeError: If prompt generation fails
        """
        self.logger.info("Starting LLM prompt generation")

        # Input validation
        if not isinstance(events_df, pd.DataFrame):
            error_msg = "Input must be a pandas DataFrame"
            self.logger.error(error_msg)
            raise ValueError(error_msg)

        if events_df.empty:
            self.logger.warning("Empty DataFrame provided for prompt generation")

        # Process events
        try:
            events_list = []
            total_events = len(events_df)

            self.logger.info(f"Processing {total_events} events for prompt...")
            for i, (_, row) in enumerate(events_df.iterrows(), 1):
                try:
                    event_dict = self._convert_datetime_fields(row.to_dict())
                    events_list.append(event_dict)
                except Exception as e:
                    self.logger.warning(
                        f"Failed to process row {i} for prompt: {str(e)}"
                    )
                    continue

            events_json = json.dumps(events_list, ensure_ascii=False, indent=2)

            template = prompt_template or self.DEFAULT_PROMPT_TEMPLATE
            prompt = template.format(events=events_json)

            self.logger.info("Successfully generated LLM prompt")
            return prompt

        except json.JSONEncodeError as e:
            error_msg = f"Failed to encode events to JSON: {str(e)}"
            self.logger.error(error_msg)
            raise RuntimeError(error_msg)
        except Exception as e:
            error_msg = f"Unexpected error during prompt generation: {str(e)}"
            self.logger.error(error_msg)
            raise RuntimeError(error_msg)
