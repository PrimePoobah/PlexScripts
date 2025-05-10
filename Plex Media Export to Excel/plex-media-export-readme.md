# Plex Media Export

A Python utility that generates a detailed Excel report of your Plex Media Server library. The tool provides an easy way to audit your media collection, highlighting video quality and tracking TV show completeness through TVMaze integration.

![License](https://img.shields.io/badge/license-GPL-blue.svg)
![Python](https://img.shields.io/badge/python-3.6+-blue.svg)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

## 🎯 Overview

This script creates a comprehensive Excel spreadsheet with two worksheets:

- **Movies**: Displays all movies with resolution-based highlighting
- **TV Shows**: Shows episode completion status verified against TVMaze

### 📋 Key Features

#### Movies Tracking

- Complete movie library inventory
- Resolution-based row highlighting:
  - 4K content (Light Green)
  - SD/480p/720p content (Yellow)
  - 1080p content (No highlight)
- Video format and file path information

#### TV Show Tracking

- Series completion overview
- Season-by-season episode verification
- Color-coded status indicators
- TVMaze database integration

## 🚀 Quick Start

### Prerequisites

- Python 3.6+
- Plex Media Server
- Plex authentication token
- Internet connection

### Installation

1. Clone the repository:

```bash
git clone https://github.com/PrimePoobah/plex-media-export.git
cd plex-media-export
```

2. Install required packages:

```bash
pip install -r requirements.txt
```

3. Configure your Plex settings:

```python
# Edit these variables in the script
PLEX_URL = 'http://{Plex_IP_or_URL}:32400'
PLEX_TOKEN = '{PlexToken}'
```

### Usage

Run the script:

```bash
python PlexMediaExport.py
```

The script will generate `PlexMediaExport_YYYYMMDD.xlsx` in your current directory.

## 📊 Output Format

### Movies Sheet

| Column           | Description               |
| ---------------- | ------------------------- |
| Title            | Movie name                |
| Video Resolution | Quality (4K, 1080p, etc.) |
| Year             | Release year              |
| Studio           | Production studio         |
| ContentRating    | Rating (PG, R, etc.)      |
| File             | Full file path            |
| Container        | File format               |

### TV Shows Sheet

| Column          | Description              |
| --------------- | ------------------------ |
| Show Title      | Series name              |
| Complete Series | Overall completion ratio |
| Season X        | Episodes present/total   |

## 🎨 Color Coding

### Movies

- 🟩 **Light Green**: 4K/UHD content
- 🟨 **Yellow**: 720p or lower
- ⬜ **No Color**: 1080p content

### TV Shows

- 🟩 **Green**: Complete series/season
- 🟥 **Red**: Incomplete series/season
- ⬛ **Gray**: Non-existent season

## ⚙️ Configuration

### Getting Your Plex Token

1. Log into Plex web interface
2. Play any media file
3. Click ⋮ (three dots menu)
4. Select "Get Info"
5. Click "View XML"
6. Look for "X-Plex-Token" in the URL

## 📝 Requirements

```plaintext
plexapi>=4.15.4
pandas>=1.3.0
openpyxl>=3.0.9
requests>=2.26.0
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 🚧 Roadmap

- [ ] Multiple Plex library support
- [ ] Command-line configuration
- [ ] Alternative TV show databases
- [ ] Additional export formats
- [ ] Enhanced logging system
- [ ] Movie database verification

## 📄 License

This project is licensed under the GPL License - see the [LICENSE](https://github.com/PrimePoobah/PlexScripts/blob/main/LICENSE) file for details.

## 🙏 Acknowledgments

- [python-plexapi](https://github.com/pkkid/python-plexapi)
- [TVMaze API](https://www.tvmaze.com/api)
- [OpenPyXL](https://openpyxl.readthedocs.io/)
- [Pandas](https://pandas.pydata.org/)

## 📮 Contact

Your Name - [@PrimePoobah](https://github.com/PrimePoobah)

Project Link: [https://github.com/PrimePoobah/PlexScripts/plex-media-export](https://github.com/PrimePoobah/PlexScripts/plex-media-export)

## ❤️ Support

If you find this project helpful, please consider giving it a ⭐️!
