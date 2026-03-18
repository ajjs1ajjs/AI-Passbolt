"""
AI-based Excel structure detector using Groq API.
Determines header row, column mapping, and data format.
"""

import json
import re
from typing import Any, Dict, List, Optional

from groq import Groq


class AIStructureDetector:
    """Detect Excel structure using AI analysis."""

    SYSTEM_PROMPT = """You are an expert at analyzing spreadsheet data structures.
Your task is to identify:
1. Which row contains headers (0-indexed)
2. Map columns to standard fields: Title, Username, Password, URL, Group, Notes
3. Detect if this is a "vertical format" where credentials are in separate rows
4. Identify group/section header rows (metadata rows that separate sections)

Respond ONLY with valid JSON in this exact format:
{
    "header_row": <number or null if no headers>,
    "columns": {
        "Title": <column_index or null>,
        "Username": <column_index or null>,
        "Password": <column_index or null>,
        "URL": <column_index or null>,
        "Group": <column_index or null>,
        "Notes": <column_index or null>
    },
    "is_vertical_format": <boolean>,
    "group_header_keywords": ["keyword1", "keyword2"],
    "confidence": <0.0 to 1.0>,
    "notes": "<brief explanation>"
}

Rules:
- Title = server/resource name (e.g., k8s-master, worker01, ingress)
- URL = IP address, hostname, or web link
- Username/Password = credentials (can be in separate columns OR combined like "user/pass")
- Group = category/folder (can be from group header rows)
- Notes = additional info, location, version, comments
- Vertical format = credentials in separate rows below server (not in same row)
- Group header rows = rows that contain section titles like "для Бобкова", "Кластре ІТ", "Demo-k8s"
- If column not present, use null
- header_row = null means data starts from row 0 without headers
- group_header_keywords = words that indicate a group header row (e.g., "для", "кластер", "group")

Common credential formats:
- Separate columns: Username in one column, Password in another
- Combined format: "username/password" in a single column (look for / separator)
- Column may be labeled "credentials", "auth", "login", or contain values like "sa/..."

Common group header patterns in Ukrainian/Russian:
- "для ..." (for ...)
- "... кластер"
- Just a name without IP (e.g., "Demo-k8s", "Sandbox")
- Rows where only first column has value and rest are empty
"""

    def __init__(
        self, api_key: Optional[str] = None, model: str = "llama-3.3-70b-versatile"
    ):
        """Initialize Groq client.

        Available models:
        - llama-3.3-70b-versatile (default, high quality)
        - llama-3.1-8b-instant (faster, less accurate)
        """
        self.api_key = api_key
        self.model = model
        self.client = None
        if api_key:
            self.client = Groq(api_key=api_key)

    def analyze(self, sample_rows: List[List[Any]], sheet_name: str = "") -> Dict:
        """
        Analyze sample rows and return structure detection.

        Args:
            sample_rows: First 10-15 rows from Excel
            sheet_name: Name of the sheet (for context)

        Returns:
            Dict with header_row, columns mapping, is_vertical_format
        """
        if not self.client:
            raise ValueError("Groq API key not provided")

        # Prepare sample for AI
        sample_text = self._format_sample(sample_rows, sheet_name)

        # Call Groq API
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": self.SYSTEM_PROMPT},
                {"role": "user", "content": sample_text},
            ],
            temperature=0.1,  # Low temperature for consistent output
            max_tokens=500,
        )

        # Parse response
        result_text = response.choices[0].message.content.strip()

        # Extract JSON from response (may have markdown code blocks)
        json_match = re.search(r"```(?:json)?\s*({.*?})\s*```", result_text, re.DOTALL)
        if json_match:
            result_text = json_match.group(1)

        try:
            result = json.loads(result_text)
        except json.JSONDecodeError as e:
            # Try to fix common issues
            result_text = result_text.replace("'", '"')
            result_text = re.sub(r",\s*}", "}", result_text)
            result_text = re.sub(r",\s*]", "]", result_text)
            try:
                result = json.loads(result_text)
            except:
                raise ValueError(
                    f"Failed to parse AI response: {e}\nResponse: {result_text}"
                )

        return result

    def _format_sample(self, sample_rows: List[List[Any]], sheet_name: str) -> str:
        """Format sample rows for AI analysis."""
        lines = []
        lines.append(f"Sheet name: {sheet_name or 'Unknown'}")
        lines.append("\nSample data (first rows):")
        lines.append("-" * 60)

        for i, row in enumerate(sample_rows[:15]):
            # Clean and truncate row data - show ALL columns (important for credentials!)
            clean_row = []
            for cell in row:  # Don't limit to 10 columns!
                if cell is None:
                    clean_row.append("")
                else:
                    cell_str = str(cell).strip()
                    # Truncate long values
                    if len(cell_str) > 50:
                        cell_str = cell_str[:47] + "..."
                    clean_row.append(cell_str)

            # Show column indices for first row
            if i == 0:
                header_indices = [f"Col{j}" for j in range(len(clean_row))]
                lines.append(f"Indices: {header_indices}")

            lines.append(f"Row {i}: {clean_row}")

        lines.append("-" * 60)
        lines.append("\nAnalyze this structure and return JSON as specified.")

        return "\n".join(lines)

    def detect_fallback(self, sample_rows: List[List[Any]]) -> Dict:
        """
        Fallback detection without AI (rule-based).
        Used when AI is unavailable or fails.
        """
        result = {
            "header_row": 0,
            "columns": {
                "Title": 0,
                "Username": None,
                "Password": None,
                "URL": 1,
                "Group": None,
                "Notes": 2,
            },
            "is_vertical_format": False,
            "confidence": 0.5,
            "notes": "Fallback detection (rule-based)",
        }

        # Check for headers in first row
        keywords = {
            "Title": ["name", "title", "server", "назва", "ім'я"],
            "Username": ["user", "login", "username", "користувач"],
            "Password": ["pass", "password", "пароль"],
            "URL": ["url", "ip", "host", "address", "link"],
            "Group": ["group", "category", "folder"],
            "Notes": ["notes", "comment", "location", "примітки"],
        }

        if sample_rows:
            first_row = [str(c).lower() if c else "" for c in sample_rows[0]]

            for col_idx, cell in enumerate(first_row):
                for field, kws in keywords.items():
                    if any(kw in cell for kw in kws):
                        result["columns"][field] = col_idx
                        result["header_row"] = 0
                        result["confidence"] = 0.7
                        break

            # Check if first row looks like data (has IP pattern)
            if len(first_row) > 1:
                if re.match(
                    r"^\d{1,3}\.\d{1,3}\.", first_row[1] if len(first_row) > 1 else ""
                ):
                    result["header_row"] = None  # No headers
                    result["columns"] = {
                        "Title": 0,
                        "URL": 1,
                        "Username": None,
                        "Password": None,
                        "Group": None,
                        "Notes": 2,
                    }

        # Check for vertical format (credentials in separate rows)
        for row in sample_rows[1:5]:
            if len(row) >= 2:
                col0 = str(row[0]).strip() if row[0] else ""
                col1 = str(row[1]).strip() if len(row) > 1 and row[1] else ""

                # Short name + password-like value (not IP)
                if (
                    len(col0) <= 15
                    and col1
                    and not re.match(r"^\d{1,3}\.", col1)
                    and not col1.startswith("http")
                ):
                    result["is_vertical_format"] = True
                    result["notes"] += " | Detected vertical format"
                    break

        return result
