# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PlexScripts is a Python 3.8+ tool that exports Plex Media Server libraries to Excel spreadsheets with advanced formatting, TVMaze integration for TV show completion tracking, customizable field selection, persistent caching, and professional logging.

**Key Features:**
- Parallel processing for both movies and TV shows
- Persistent TVMaze cache (90-95% faster on repeat runs)
- Rotating log files with timestamps
- Retry logic with exponential backoff for network resilience
- Environment variable validation
- Dictionary-based field processing for maintainability

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

### Logging System (Lines 149-183)
Professional logging infrastructure with dual output:
- **File logging**: Rotating log files in `logs/` directory (5MB max, 5 backups)
- **Console logging**: INFO level for user feedback
- **Log format**: `YYYY-MM-DD HH:MM:SS - LEVEL - message`
- All operations logged: connections, API calls, errors, warnings

### Environment Validation (Lines 186-214)
Validates configuration at startup before processing begins:
- Checks `PLEX_URL` format (must start with http:// or https://)
- Validates `PLEX_TOKEN` presence and minimum length
- Validates field selections against known fields
- Provides helpful error messages and warnings

### Retry Logic (Lines 217-250)
Decorator `@retry_on_failure()` with exponential backoff:
- Default: 3 retries with 1.0s initial delay, 2x backoff
- Handles `requests.exceptions.RequestException`
- Applied to TVMaze API calls for network resilience
- Logs retry attempts and failures

### Persistent Cache System (Lines 255-369)
TVMaze API responses cached to disk for 30 days:
- **Cache file**: `.tvmaze_cache.pkl` (pickle format)
- **Structure**: Versioned dict with timestamps per entry
- **Expiration**: Automatic cleanup of entries >30 days old
- **Performance**: 90-95% faster on subsequent runs
- **Functions**:
  - `load_tvmaze_cache()`: Loads and validates cache on startup
  - `save_tvmaze_cache()`: Persists cache on shutdown
  - `get_from_cache()` / `add_to_cache()`: Cache access helpers

### Configuration System (Lines 49-122)
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

#### TV Show Processing (Lines 469-589)
1. **`process_show_metadata(show_obj)`**: Extracts base show metadata using dictionary-based field processors
2. **`process_single_show(show_obj)`**: Processes a single TV show (designed for parallel execution)
   - Attempts IMDB ID lookup first, falls back to title search
   - Uses `s.leafCount` for O(1) episode counts (faster than `len(s.episodes())`)
   - Combines Plex season/episode counts with TVMaze expected counts
3. **`get_show_details(shows)`**: **Parallel processing** using ThreadPoolExecutor
   - Worker count: `min(10, len(shows))`
   - Uses `executor.map()` for concurrent show processing
   - Returns both processed data and `max_seasons_overall` for column generation
   - Significantly faster than sequential processing (5-10x speedup)

### TVMaze Integration (Lines 458-564)

**Public Function**: `get_tvmaze_show_info(search_term, is_imdb_id=False)`
- Checks **persistent cache first** (`.tvmaze_cache.pkl`)
- Falls back to API call if not cached
- Automatically caches API results for future runs
- Cache key format: `"imdb:{id}"` or `"title:{name}"`

**Internal Function**: `_fetch_tvmaze_show_info_from_api(search_term, is_imdb_id=False)`
- Decorated with `@retry_on_failure()` for network resilience
- Two search strategies:
  1. IMDB ID lookup: `/lookup/shows?imdb={id}` (more reliable)
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

### Field Processing Architecture (Dictionary-Based)
The script uses dictionary-based field processors for maintainability:
- **`_get_movie_field_processors()`** (Lines 308-343): Returns dict mapping field names to lambda extractors
- **`_get_show_field_processors()`** (Lines 346-368): Returns dict mapping field names to lambda extractors
- Each processor takes the Plex object (and media/parts for movies) and returns the value
- Adding new fields: Simply add an entry to the appropriate processor dictionary
- Benefits: Eliminates long if-elif chains, easier to maintain and extend

When adding new fields:
1. Add field name to `ALL_POSSIBLE_MOVIE_FIELDS` or `ALL_POSSIBLE_SHOW_FIELDS`
2. Add entry to `_get_movie_field_processors()` or `_get_show_field_processors()`
3. Use `getattr()` with defaults, handle None values gracefully
4. Default to 'N/A' for missing data, or 0 for count fields

### TVMaze API Considerations
- **Persistent cache** dramatically reduces API calls (90-95% reduction on subsequent runs)
- Cache entries expire after 30 days, balancing freshness with performance
- IMDB ID lookup is more reliable than title search when available
- Some shows may not exist in TVMaze database; always handle `None` returns
- Episodes without season numbers (e.g., some specials) are intentionally filtered out
- **Retry logic** automatically handles transient network failures (3 retries with exponential backoff)

### Performance Characteristics
- **First run**: Builds cache, same speed as original
- **Subsequent runs**: 90-95% faster for TV show processing
- **Parallel processing**: 5-10x faster than sequential for large libraries
- **Optimized Plex API calls**: Uses `leafCount` instead of fetching all episodes

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
