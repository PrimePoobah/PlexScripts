# Plex Media Export Tools

![GPL License](https://img.shields.io/badge/license-GPL--3.0-blue)
![Python Version](https://img.shields.io/badge/python-3.6%2B-blue)
![Plex API](https://img.shields.io/badge/PlexAPI-compatible-brightgreen)
![Excel Output](https://img.shields.io/badge/Excel-report-success)
![TVMaze Integration](https://img.shields.io/badge/TVMaze-integrated-informational)
![Pandas](https://img.shields.io/badge/Pandas-powered-ff69b4)
![Maintenance](https://img.shields.io/badge/maintained-yes-green.svg)
![Dotenv Support](https://img.shields.io/badge/dotenv-configurable-orange)
![Multi-Threading](https://img.shields.io/badge/processing-parallel-blue)

A comprehensive collection of Python utilities designed to export content from your Plex Media Server into detailed Excel reports. These tools help you catalog, audit, and manage your movie and TV show collections with powerful features like resolution tracking and TV show completeness verification.

## 📋 Overview

This repository contains three specialized scripts for Plex content management:

1. [**Plex Media Export**](https://github.com/PrimePoobah/PlexScripts/tree/main/Plex%20Media%20Export%20to%20Excel) - A complete solution that combines movie and TV show tracking in a single Excel workbook
2. [**Plex Movie List Exporter**](https://github.com/PrimePoobah/PlexScripts/tree/main/Plex%20Movie%20List%20Export%20to%20Excel) - A focused utility for movie library cataloging
3. [**Plex TV Show Audit Tool**](https://github.com/PrimePoobah/PlexScripts/tree/main/Plex%20TV%20Show%20Export%20to%20Excel) - A specialized script for tracking TV show completion status with TVMaze integration

All tools are designed to be user-friendly, performance-optimized, and provide valuable insights into your media collection.

## ✨ Key Features

### 🎬 Movie Tracking

- Complete inventory of your movie library
- Resolution-based highlighting (4K, 1080p, 720p, SD)
- Technical details (container format, file path)
- Content metadata (release year, studio, rating)
- Customizable export fields via .env configuration
- Alphabetical sorting for easy reference

### 📺 TV Show Tracking

- Series completion overview with TVMaze verification
- Season-by-season episode counting
- Color-coded status indicators:
  - 🟩 **Green**: Complete series/seasons
  - 🟥 **Red**: Incomplete series/seasons
  - ⬛ **Gray**: Non-existent seasons
- Missing episode identification
- Customizable TV show metadata fields

### 🛠️ Advanced Features

- Multi-threaded processing for improved performance
- Environment variables from .env file for simple configuration
- Memory-optimized Excel generation for large libraries
- Cached TVMaze lookups to reduce API calls
- Detailed progress reporting during execution
- Error handling for missing or corrupt media files
- Sortable and filterable Excel tables
- Customizable output directory

## 🚀 Getting Started

### System Requirements

- Python 3.6 or higher
- Access to a Plex Media Server
- Plex authentication token
- Internet connection (for TVMaze integration)

### Installation

1. Clone the repository:

```bash
git clone https://github.com/PrimePoobah/plex-media-export.git
cd plex-media-export
```

2. Install required packages:

```bash
pip install plexapi pandas openpyxl requests python-dotenv
```

3. Create and configure your .env file:

```bash
# Copy the example .env file
cp .env.example .env
# Edit with your preferred text editor
nano .env
```

4. Configure your Plex settings in the .env file:

```
PLEX_URL="http://{Plex_IP_or_URL}:32400"
PLEX_TOKEN="{YourPlexToken}"
PLEX_EXPORT_DIR="{OptionalCustomExportPath}"
PLEX_MOVIE_EXPORT_FIELDS="Title,Year,Studio,ContentRating,Video Resolution,File Path,Container,Duration (min)"
PLEX_SHOW_EXPORT_FIELDS="Title,Year,Studio,ContentRating"
```

### Finding Your Plex Token

1. Log into the Plex web interface
2. Play any media file
3. Click the three dots menu (⋮)
4. Select "Get Info"
5. Click "View XML"
6. Look for "X-Plex-Token" in the URL

## 🧰 Script Details

### PlexMediaExport.py

The most comprehensive script that combines movie and TV show tracking in a single Excel workbook with extensive customization options via .env configuration.

```bash
python PlexMediaExport.py
```

Output: `PlexMediaExport_YYYYMMDD_HHMMSS.xlsx` with one worksheet per library section:

- **Movies**: Complete movie library with resolution highlighting
- **TV Shows**: Series completion status with TVMaze verification

### plex_movie_export.py

A focused script for movie library cataloging.

```bash
python plex_movie_export.py
```

Output: `plex_movies.xlsx` containing your complete movie library details.

### plex_tv_shows.py

A specialized script for TV show completion tracking.

```bash
python plex_tv_shows.py
```

Output: `plex_tv_shows_YYYYMMDD.xlsx` showing series and season completion status.

## 📊 Excel Report Details

### Movies Worksheet Format

| Column              | Description               |
| ------------------- | ------------------------- |
| Title               | Movie name                |
| Video Resolution    | Quality (4K, 1080p, etc.) |
| Year                | Release year              |
| Studio              | Production studio         |
| ContentRating       | Rating (PG, R, etc.)      |
| File                | Full file path            |
| Container           | File format               |
| Duration (min)      | Movie length in minutes   |
| _Additional fields_ | Customizable via .env     |

### TV Shows Worksheet Format

| Column              | Description              |
| ------------------- | ------------------------ |
| Show Title          | Series name              |
| Complete Series     | Overall completion ratio |
| Season X            | Episodes present/total   |
| _Additional fields_ | Customizable via .env    |

## 🎨 Color Coding

### Movies

- 🟩 **Light Green**: 4K/UHD content
- 🟨 **Yellow**: 720p or lower resolution
- ⬜ **No Color**: 1080p content (standard)

### TV Shows

- 🟩 **Green**: Complete series/season
- 🟥 **Red**: Incomplete series/season
- ⬛ **Gray**: Non-existent season

## 🔧 Customization

The script supports extensive customization through the .env file:

### Available Movie Fields

```
Title, Year, Studio, ContentRating, Video Resolution,
Bitrate (kbps), File Path, Container, Duration (min),
AddedAt, LastViewedAt, OriginallyAvailableAt, Summary, Tagline,
AudienceRating, Rating, Collections, Genres, Labels,
AspectRatio, AudioChannels, AudioCodec, VideoCodec, VideoFrameRate,
Height, Width, ViewCount, SkipCount
```

### Available TV Show Fields

```
Title, Year, Studio, ContentRating, Summary, Tagline,
AddedAt, LastViewedAt, OriginallyAvailableAt,
AudienceRating, Rating, Collections, Genres, Labels,
ViewCount, SkipCount
```

## 📝 Requirements

```
plexapi>=4.15.4
pandas>=1.3.0
openpyxl>=3.0.9
requests>=2.26.0
python-dotenv>=0.19.0
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the GNU Affero General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [python-plexapi](https://github.com/pkkid/python-plexapi) for the Plex API integration
- [TVMaze API](https://www.tvmaze.com/api) for TV show database information
- [OpenPyXL](https://openpyxl.readthedocs.io/) for Excel file generation
- [Pandas](https://pandas.pydata.org/) for data processing
- [Plex](https://www.plex.tv/) for their amazing media server platform
- [nledenyi](https://github.com/nledenyi) for contributions to the project
- [python-dotenv](https://github.com/theskumar/python-dotenv) for environment variable management

## 📮 Contact

PrimePoobah - [@PrimePoobah](https://github.com/PrimePoobah)

Project Link: [https://github.com/PrimePoobah/plex-media-export](https://github.com/PrimePoobah/plex-media-export)

## ❓ FAQ

### Q: Can I run these scripts on a headless server?

A: Yes, all scripts are command-line based and don't require a GUI.

### Q: Will these scripts modify my Plex library?

A: No, they only read data from your Plex server and don't make any changes.

### Q: How often should I run these exports?

A: It depends on how frequently you add content. Weekly or monthly is typical.

### Q: Can I customize the Excel formatting?

A: Yes, you can modify the styling variables in the scripts to customize colors and formats.

### Q: Can I choose which fields to export for movies and TV shows?

A: Yes, you can specify the exact fields you want in your .env file using the PLEX_MOVIE_EXPORT_FIELDS and PLEX_SHOW_EXPORT_FIELDS variables.

### Q: Why do some shows not appear in the TV Show report?

A: This usually happens when a show name doesn't match between Plex and TVMaze.

### Q: Where are the export files saved?

A: By default, they're saved in the same directory as the script. You can specify a custom directory with the PLEX_EXPORT_DIR variable in your .env file.

## ❤️ Support

If you find these tools useful, please consider:

- Giving this project a ⭐️ on GitHub
- Sharing it with other Plex users
- Contributing improvements back to the project
