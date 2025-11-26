# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PlexScripts is a Python tool that exports Plex Media Server libraries to Excel spreadsheets with advanced formatting, TVMaze integration for TV show completion tracking, and customizable field selection.

## Essential Commands

### Setup and Installation
```bash
# Install dependencies
pip install plexapi pandas openpyxl requests python-dotenv

# Configure environment
cp .env.example .env
# Edit .env with your PLEX_URL and PLEX_TOKEN
```

### Running the Script
```bash
# Main export script (exports both movies and TV shows)
python PlexMediaExport.py
```

### Finding Your Plex Token
1. Log into Plex web interface
2. Play any media file
3. Click the three dots (⋮) → "Get Info" → "View XML"
4. Find "X-Plex-Token" in the URL

## High-Level Architecture

### Configuration System (Lines 40-122)
The script uses a sophisticated field configuration system controlled via `.env` file:
- `PLEX_MOVIE_EXPORT_FIELDS` and `PLEX_SHOW_EXPORT_FIELDS` define which metadata fields to export
- Fields are validated against `ALL_POSSIBLE_MOVIE_FIELDS` and `ALL_POSSIBLE_SHOW_FIELDS`
- Invalid or missing field configurations fall back to sensible defaults
- This system allows users to customize exports without code changes

### Processing Flow

#### Movie Processing (Lines 303-426)
1. **`process_movie(movie)`**: Extracts only requested fields from Plex movie objects
   - Uses conditional attribute access to avoid unnecessary API calls
   - Only fetches media/parts objects if media-dependent fields are requested
   - Handles missing attributes gracefully with 'N/A' defaults
2. **`get_movie_details(movies)`**: Parallel processing using ThreadPoolExecutor
   - Worker count: `min(10, (os.cpu_count() or 1) + 4)`
   - Significantly improves performance for large libraries

#### TV Show Processing (Lines 427-544)
1. **`process_show_metadata(show_obj)`**: Extracts base show metadata
2. **`get_show_details(shows)`**: Sequential processing with TVMaze integration
   - Attempts IMDB ID lookup first, falls back to title search
   - Combines Plex season/episode counts with TVMaze expected counts
   - Returns both processed data and `max_seasons_overall` for column generation

### TVMaze Integration (Lines 186-257)

**Key Function**: `get_tvmaze_show_info(search_term, is_imdb_id=False)`
- Decorated with `@lru_cache(maxsize=256)` to avoid redundant API calls
- Two search strategies:
  1. IMDB ID lookup: `/lookup/shows?imdb={id}`
  2. Title search: `/search/shows?q={title}` (uses first result)
- Returns structured data: `{'total_seasons': int, 'seasons': {season_num: {'total_episodes': int}}}`
- Filters out episodes without season numbers (some specials)
- Robust error handling for network issues and missing data

### Excel Generation Strategy

#### Movie Worksheet (Lines 645-739)
- **Resolution-based color coding** (Lines 691-714):
  - 4K/UHD/2160p → Light Green (`#77b190`)
  - 1080p → White
  - 720p or lower → Yellow
- Uses Pandas DataFrame for sorting and data manipulation
- Sorts by 'Title' if selected, otherwise first selected field
- Applies `TableStyleMedium9` with row stripes

#### TV Show Worksheet (Lines 741-863)
- **Dynamic column generation**: Base fields + "Series Complete" + Season columns (S00, S01, etc.)
- **Freeze panes** after base metadata columns for easier horizontal scrolling
- **Completion tracking colors**:
  - Green: Season/series complete (Plex episodes ≥ TVMaze expected)
  - Red: Incomplete (some episodes present, but < expected)
  - Yellow: Unknown TVMaze data (shows `X/?`)
  - Gray: Season doesn't exist per TVMaze
- **Pre-sorting** (Lines 757-759): Shows are sorted before row generation to ensure Excel order matches sort
- S00 (Specials) column only appears if any show has specials in Plex

### Styling System (Lines 129-152, 546-566)
Centralized `STYLES` dictionary contains all PatternFills, Borders, Fonts, and Alignments:
- Promotes consistency across worksheets
- `apply_cell_styling()` function applies border, alignment, font, and fill in one call
- Summary/Tagline fields use `wrap_text=True` alignment

### Session Management
- **Plex**: Dedicated `Session()` created in `connect_to_plex()` for PlexAPI calls
- **TVMaze**: Global `session = Session()` (line 127) provides connection pooling for API requests
- Both improve performance through connection reuse

## Important Implementation Notes

### Field Processing Optimization
When adding or modifying field extraction in `process_movie()` or `process_show_metadata()`:
- Always check if field is in `SELECTED_*_FIELDS` before processing
- For media-dependent movie fields, ensure `media` object is available (Lines 320-328)
- Use `hasattr()` and null checks before accessing Plex object attributes
- Default to 'N/A' for missing data, or 0 for count fields

### TVMaze API Considerations
- Rate limits exist but are generous; caching via `@lru_cache` prevents most issues
- IMDB ID lookup is more reliable than title search when available
- Some shows may not exist in TVMaze database; always handle `None` returns
- Episodes without season numbers (e.g., some specials) are intentionally filtered out

### Excel Table Naming
- Table names must be sanitized: no spaces, must start with a letter (Lines 584-586)
- Sheet names are limited to 31 characters and have character restrictions (Line 659)
- Always use `create_table()` helper function to ensure proper formatting

### Color Fill Debugging
- Resolution matching is case-insensitive and handles variations (e.g., "1080p", "1080")
- Debug print statements are preserved but commented out (Lines 686-721)
- Table row stripes (`showRowStripes=True`) can interfere with custom fills; current implementation accepts this trade-off

## Environment Variables

Required:
- `PLEX_URL`: Plex server URL (e.g., "http://192.168.1.100:32400")
- `PLEX_TOKEN`: Plex authentication token

Optional:
- `PLEX_EXPORT_DIR`: Output directory (defaults to current directory)
- `PLEX_MOVIE_EXPORT_FIELDS`: Comma-separated list of movie fields to export
- `PLEX_SHOW_EXPORT_FIELDS`: Comma-separated list of TV show base fields to export

## Output

The script generates a timestamped Excel file: `PlexMediaExport_YYYYMMDD_HHMMSS.xlsx`
- One worksheet per Plex library section (movie or TV show)
- Sortable, filterable Excel tables with consistent styling
- Color-coded cells provide visual completion status
