# 📦 PlexScripts - Export Your Plex Library to Excel

![GPL License](https://img.shields.io/badge/license-GPL--3.0-blue)
![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![Plex API](https://img.shields.io/badge/PlexAPI-compatible-brightgreen)
![Excel Output](https://img.shields.io/badge/Excel-report-success)
![TVMaze Integration](https://img.shields.io/badge/TVMaze-integrated-informational)
![Pandas](https://img.shields.io/badge/Pandas-powered-ff69b4)
![Maintenance](https://img.shields.io/badge/maintained-yes-green.svg)
![Dotenv Support](https://img.shields.io/badge/dotenv-configurable-orange)
![Multi-Threading](https://img.shields.io/badge/processing-parallel-blue)

## Overview

A Python tool that exports your Plex Media Server library to professionally formatted Excel spreadsheets. Track your movies with resolution-based color coding and your TV shows with TVMaze integration for completion tracking.

---

## 🎞️ What Does This Do?

**PlexMediaExport.py** - A comprehensive export script that processes both movies and TV shows from your Plex server into a single timestamped Excel file with:

- **Movies**: Detailed metadata with resolution-based color highlighting
- **TV Shows**: Season-by-season completion tracking powered by TVMaze API

---

## Features

### 🎬 Movie Export
- Complete metadata inventory with customizable field selection
- Resolution-based color highlighting:
  - 🟩 4K/UHD/2160p: Light green
  - ⬜ 1080p: White
  - 🟨 720p or lower: Yellow
- Technical details: codecs, bitrate, container, aspect ratio, audio channels
- Content metadata: ratings, genres, collections, labels, summaries
- File paths and viewing statistics
- Alphabetically sorted by title (or first selected field)

### 📺 TV Show Tracking
- TVMaze API integration for accurate episode counts
- Series completion status tracking
- Season-by-season breakdown with episode counts
- Color-coded status indicators:
  - 🟩 Green: Complete (Plex episodes ≥ TVMaze expected)
  - 🟥 Red: Incomplete (some episodes present, but < expected)
  - 🟨 Yellow: Unknown TVMaze data (shows X/?)
  - ⬛ Gray: Season doesn't exist per TVMaze
- Handles specials (Season 0) when present

### 💥 Performance & Reliability
- **Parallel processing**: Multi-threaded processing for both movies and TV shows
- **Persistent TVMaze cache**: 90-95% faster on repeat runs (30-day cache lifetime)
- **Professional logging**: Rotating log files with timestamps in `logs/` directory
- **Retry logic**: Exponential backoff for network resilience
- **Environment validation**: Catches configuration errors before processing
- **Memory-optimized**: Efficient Excel generation for large libraries
- **Sortable Excel tables**: Professional formatting with TableStyleMedium9 style

---

## Installation

### Requirements
- Python 3.8 or higher
- Plex Media Server with authentication token
- Internet connection (for TVMaze API lookups)

### Setup

```bash
git clone https://github.com/PrimePoobah/PlexScripts.git
cd PlexScripts
pip install plexapi pandas openpyxl requests python-dotenv
```

### Configuration

1. Copy the example environment file:
```bash
cp .env.example .env
```

2. Edit `.env` with your Plex credentials:
```env
PLEX_URL="http://YOUR_PLEX_IP:32400"
PLEX_TOKEN="YOUR_PLEX_TOKEN_HERE"
PLEX_EXPORT_DIR=""  # Optional: specify output directory
PLEX_MOVIE_EXPORT_FIELDS="Title,Year,Studio,ContentRating,Video Resolution,File Path,Container,Duration (min)"
PLEX_SHOW_EXPORT_FIELDS="Title,Year,Studio,ContentRating"
```

### Finding Your Plex Token
1. Log into Plex web interface
2. Play any media file
3. Click the three dots (⋮) → "Get Info"
4. Click "View XML"
5. Find "X-Plex-Token" in the URL parameter

---

## Usage

### Running the Script

```bash
python PlexMediaExport.py
```

The script will:
1. Connect to your Plex server
2. Process all movie and TV show libraries
3. Fetch TVMaze data for TV shows (with persistent caching)
4. Generate a timestamped Excel file: `PlexMediaExport_YYYYMMDD_HHMMSS.xlsx`

Each library section gets its own worksheet in the output file.

### First Run vs. Subsequent Runs
- **First run**: Builds TVMaze cache, processes at normal speed
- **Subsequent runs**: 90-95% faster for TV shows using cached data
- Cache file: `.tvmaze_cache.pkl` (auto-expires entries after 30 days)

---

## Output Format

### Movie Worksheet Columns

All selected fields are included (customizable via `.env`). Default fields:

| Column            | Description                              |
|-------------------|------------------------------------------|
| Title             | Movie title                              |
| Year              | Release year                             |
| Studio            | Production studio                        |
| ContentRating     | MPAA/content rating                      |
| Video Resolution  | Resolution with color coding             |
| File Path         | Full file path on server                 |
| Container         | File container format (mkv, mp4, etc.)   |
| Duration (min)    | Runtime in minutes                       |

**Available fields**: Title, Year, Studio, ContentRating, Video Resolution, Bitrate (kbps), File Path, Container, Duration (min), AddedAt, LastViewedAt, OriginallyAvailableAt, Summary, Tagline, AudienceRating, Rating, Collections, Genres, Labels, AspectRatio, AudioChannels, AudioCodec, VideoCodec, VideoFrameRate, Height, Width, ViewCount, SkipCount

### TV Show Worksheet Columns

Base metadata columns (customizable) + completion tracking:

| Column                       | Description                          |
|------------------------------|--------------------------------------|
| Selected base fields         | Title, Year, Studio, etc.            |
| Series Complete (Plex/TVMaze)| Overall completion ratio (X/Y)       |
| S00, S01, S02...             | Per-season episode counts (X/Y)      |

**Available base fields**: Title, Year, Studio, ContentRating, Summary, Tagline, AddedAt, LastViewedAt, OriginallyAvailableAt, AudienceRating, Rating, Collections, Genres, Labels, ViewCount, SkipCount

---

## Customization

### Customizing Export Fields

Edit your `.env` file to select which fields to export:

**Movie Fields Example**:
```env
PLEX_MOVIE_EXPORT_FIELDS="Title,Year,Video Resolution,File Path,Genres,Collections"
```

**TV Show Fields Example**:
```env
PLEX_SHOW_EXPORT_FIELDS="Title,Year,Summary,Genres"
```

Leave blank or comment out to use defaults.

### Output Directory

Specify a custom output directory:
```env
PLEX_EXPORT_DIR="/path/to/export/directory"
```

Defaults to current directory if not specified.

---

## Dependencies

```
plexapi>=4.15.4
pandas>=1.3.0
openpyxl>=3.0.9
requests>=2.26.0
python-dotenv>=0.19.0
```

Install all at once:
```bash
pip install plexapi pandas openpyxl requests python-dotenv
```

---

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/YourFeature`)
3. Commit your changes (`git commit -m 'Add YourFeature'`)
4. Push to the branch (`git push origin feature/YourFeature`)
5. Open a Pull Request

---

## License

This project is licensed under the **GNU AGPL v3.0** - see the [LICENSE](LICENSE) file for details.

---

## Acknowledgments

- [python-plexapi](https://github.com/pkkid/python-plexapi) - Plex API integration
- [TVMaze API](https://www.tvmaze.com/api) - TV show episode data
- [OpenPyXL](https://openpyxl.readthedocs.io/) - Excel file generation
- [Pandas](https://pandas.pydata.org/) - Data processing
- [python-dotenv](https://github.com/theskumar/python-dotenv) - Environment variable management
- [nledenyi](https://github.com/nledenyi) - Contributor

---

## Project Links

**Author**: [PrimePoobah](https://github.com/PrimePoobah)
**Repository**: [PlexScripts](https://github.com/PrimePoobah/PlexScripts)

---

## Troubleshooting

### Logging

The script creates detailed log files in the `logs/` directory:
- **File format**: `plex_export_YYYYMMDD.log`
- **Rotation**: Automatically rotates at 5MB, keeps last 5 files
- **Contents**: Connections, API calls, errors, warnings, debug info

**Checking logs**:
```bash
# View today's log (Linux/macOS)
cat logs/plex_export_$(date +%Y%m%d).log

# Watch logs in real-time
tail -f logs/plex_export_*.log

# Search for errors
grep "ERROR" logs/*.log
```

### Cache Management

**TVMaze Cache** (`.tvmaze_cache.pkl`):
- Location: Export directory (same as output files)
- Lifetime: 30 days (auto-expires old entries)
- Can be safely deleted - rebuilds on next run (slower initial processing)

---

## FAQ

### Will this work on a headless server?
Yes, the script runs entirely in the terminal with no GUI required.

### Will this modify my Plex library?
No, the script is read-only. It only queries data from your Plex server and TVMaze API.

### How often should I run this?
As often as you want updated exports. The persistent cache makes subsequent runs very fast (90-95% faster for TV shows).

### Can I customize the Excel formatting/colors?
Yes, edit the `STYLES` dictionary and color logic in [PlexMediaExport.py](PlexMediaExport.py). See [CLAUDE.md](CLAUDE.md) for architecture details.

### Why don't some TV shows have TVMaze data?
Some shows may not exist in the TVMaze database, or the title/IMDB matching may fail. The script shows `X/?` (yellow) for unknown data.

### Where does the output file go?
By default, the current directory. Set `PLEX_EXPORT_DIR` in `.env` to specify a different location.

### Can I export only movies or only TV shows?
Currently, the script processes all sections. You can manually delete unwanted worksheets from the Excel file, or modify the code to filter sections.

---

## Support

If you find this tool helpful:
- ⭐ Star the repository
- Report issues on [GitHub Issues](https://github.com/PrimePoobah/PlexScripts/issues)
- Contribute improvements via Pull Requests
