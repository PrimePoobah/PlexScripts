# 📦 PlexScripts - Because Your Digital Media Hoard Deserves Excel-Level OCD

![GPL License](https://img.shields.io/badge/license-GPL--3.0-blue)
![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![Plex API](https://img.shields.io/badge/PlexAPI-compatible-brightgreen)
![Excel Output](https://img.shields.io/badge/Excel-report-success)
![TVMaze Integration](https://img.shields.io/badge/TVMaze-integrated-informational)
![Pandas](https://img.shields.io/badge/Pandas-powered-ff69b4)
![Maintenance](https://img.shields.io/badge/maintained-yes-green.svg)
![Dotenv Support](https://img.shields.io/badge/dotenv-configurable-orange)
![Multi-Threading](https://img.shields.io/badge/processing-parallel-blue)

## Hey There, Beautiful Media Hoarders! 👋

So you've got a Plex server stuffed with more cinematic gold and TV trash than a Blockbuster from 2003 (RIP, sweet prince), and you want to know what the hell you actually have? Well buckle up, buttercup!

This is a Python-powered bad boy that'll rip your Plex library into Excel spreadsheets so detailed, they'd make the IRS weep tears of joy. We're talking movie catalogs, TV show completeness tracking, and enough color-coding to make a rainbow jealous.

---

## 🎞️ What Is This Sorcery?

**PlexMediaExport.py** - The Swiss Army Chainsaw of Excel reports. One script to rule them all. Movies AND TV shows. Together. Harmony. Beautiful, beautiful harmony.

- **Movies**: Detailed metadata with resolution-based color highlighting (shame that 480p collection!)
- **TV Shows**: Season-by-season completion tracking powered by TVMaze API (because manual counting is for psychopaths)

---

## 🧠 Features That Punch Mediocrity in the Face

### 🎬 For Movie Maniacs (Or: "How Many Times Did I Buy The Same Movie?")
- Complete inventory because you probably forgot you own Green Lantern (we don't talk about Green Lantern)
- Resolution-based highlighting that'll shame you for your 480p collection:
  - 🟩 4K/UHD/2160p → Light green (oh shiny!)
  - ⬜ 1080p → White (the standard bearer of mediocrity)
  - 🟨 720p or lower → Yellow (we don't judge... much)
- Technical details for the nerds: codecs, bitrate, container, aspect ratio, audio channels
- Content metadata so you can remember why you bought Cats (spoiler: there's no good reason)
- Customizable fields via .env because I'm not your mom – configure it yourself
- Alphabetically sorted because chaos is only fun in combat
- It's like IMDB got Excel-itis

### 📺 TV Show Tracking (Or: "The 'Do I Actually Have All the Episodes?' Tool")
- Finds missing episodes you didn't know you were avoiding
- Series completion tracking with TVMaze integration (because manual counting is for psychopaths)
- Season-by-season breakdown that's more detailed than my therapy sessions
- Color-coded status indicators:
  - 🟩 Green: Complete (like my collection of regrets)
  - 🟥 Red: Incomplete (like my understanding of healthy relationships)
  - 🟨 Yellow: Unknown TVMaze data (shows X/? - existential crisis time!)
  - ⬛ Gray: This never existed (like my social life)
- Handles specials (Season 0) because some shows are extra like that

### 💥 Advanced Features (Because We're Fancy Like That)
- **Multi-threaded processing** (faster than my mouth in a fight) - Parallel movie AND TV show processing
- **Persistent TVMaze cache** (90-95% faster on repeat runs) - Cached lookups survive between runs like cockroaches
- **Professional logging** (debugging without the headache) - Rotating log files with timestamps in `logs/` directory
- **Retry logic with exponential backoff** (network hiccups don't phase us - we're resilient like that)
- **Environment variable validation** (catches config errors before you waste time)
- **Environment variables from .env files** (because hardcoding is what killed the dinosaurs)
- **Memory-optimized Excel generation** (won't crash your potato computer)
- **Sortable Excel tables** (organization is my kink... I mean, my thing)

---

## 🚀 Installation Instructions (or "How to Not Screw It Up")

### 🧰 Requirements
- Python 3.8+ (we're modern now, grandpa)
- Plex server (not imaginary)
- Plex token (like a golden key but nerdier)
- Internet (for TVMaze, not for my OnlyFans)

### 🧙‍♂️ Summon the Tools

```bash
git clone https://github.com/PrimePoobah/PlexScripts.git
cd PlexScripts
pip install plexapi pandas openpyxl requests python-dotenv
```

### ⚙️ Configuration (This Is Where the Magic Happens)

**Step 1:** Copy the example environment file:
```bash
cp .env.example .env
nano .env  # Or edit in Notepad if you're stuck in 2005
```

**Step 2:** Fill in your actual Plex credentials:
```env
PLEX_URL="http://YOUR_PLEX_IP:32400"
PLEX_TOKEN="YOUR_PLEX_TOKEN_HERE"
PLEX_EXPORT_DIR=""  # Optional: where you want your files (defaults to current dir)
PLEX_MOVIE_EXPORT_FIELDS="Title,Year,Studio,ContentRating,Video Resolution,File Path,Container,Duration (min)"
PLEX_SHOW_EXPORT_FIELDS="Title,Year,Studio,ContentRating"
```

### 🎟️ The Great PLEX_TOKEN Hunt
1. Log into Plex web interface (you know, that thing you use to procrastinate)
2. Play any media file (I suggest *Deadpool*)
3. Click the three dots (⋮) because apparently we're all ancient Greeks now
4. Select "Get Info" (getting philosophical, are we?)
5. Click "View XML" (now we're speaking in tongues)
6. Find "X-Plex-Token" in the URL (congratulations, you're now a hacker)

---

## 🎯 Running This Bad Boy

### Fire It Up!

This script does everything. It's like the Swiss Army knife of Plex tools, if Swiss Army knives could judge your media collection.

```bash
python PlexMediaExport.py
```

**What happens next:**
1. Connects to your Plex server (fingers crossed!)
2. Processes ALL your movie and TV show libraries (buckle up)
3. Fetches TVMaze data for TV shows (with caching because we're not savages)
4. Spits out a timestamped Excel file: `PlexMediaExport_YYYYMMDD_HHMMSS.xlsx`

*Because timestamps are like signatures – they prove you were here*

Each library section gets its own worksheet. Movies here, TV shows there. Everything organized like Marie Kondo's sock drawer.

### 🏃 First Run vs. Subsequent Runs (The Need for Speed)
- **First run**: Builds TVMaze cache, processes at normal speed (grab a coffee)
- **Subsequent runs**: 90-95% FASTER for TV shows using cached data (grab a shot)
- **Cache file**: `.tvmaze_cache.pkl` (auto-expires after 30 days like old milk, but less smelly)

---

## 📊 Spreadsheet Glory Details

### Movies Sheet (The "Do I Really Own This?" Tab)

All selected fields are included (customizable via `.env` because you're special). Default fields:

| Column            | Description                              |
|-------------------|------------------------------------------|
| Title             | Movie name                               |
| Year              | Release year                             |
| Studio            | Studio magic                             |
| ContentRating     | MPAA rating (or "oops")                  |
| Video Resolution  | 4K? Fancy. SD? Eww.                      |
| File Path         | Full path to shame                       |
| Container         | MKV? MP4? VHS? (please no)               |
| Duration (min)    | Minutes of regret/joy                    |

**The Full Buffet of Available Fields**: Title, Year, Studio, ContentRating, Video Resolution, Bitrate (kbps), File Path, Container, Duration (min), AddedAt, LastViewedAt, OriginallyAvailableAt, Summary, Tagline, AudienceRating, Rating, Collections, Genres, Labels, AspectRatio, AudioChannels, AudioCodec, VideoCodec, VideoFrameRate, Height, Width, ViewCount, SkipCount

*Mix and match via your `.env` file - you decide what matters!*

### TV Shows Sheet (The "Series Complete?" Guilt Trip)

Base metadata columns (your choice) + completion tracking columns (my gift to you):

| Column                       | Description                          |
|------------------------------|--------------------------------------|
| Selected base fields         | Title, Year, Studio, etc.            |
| Series Complete (Plex/TVMaze)| Overall completion ratio (X/Y)       |
| S00, S01, S02...             | Per-season episode counts (X/Y)      |

**Available Base Fields**: Title, Year, Studio, ContentRating, Summary, Tagline, AddedAt, LastViewedAt, OriginallyAvailableAt, AudienceRating, Rating, Collections, Genres, Labels, ViewCount, SkipCount

---

## 🔧 Customization (Make It Your Own, Like A Bad Tattoo)

### Customizing Export Fields

Edit your `.env` file like a wizard:

**Movie Fields Example (Choose Your Destiny)**:
```env
PLEX_MOVIE_EXPORT_FIELDS="Title,Year,Video Resolution,File Path,Genres,Collections"
```

**TV Show Fields Example (Choose Your Own Binge-venture)**:
```env
PLEX_SHOW_EXPORT_FIELDS="Title,Year,Summary,Genres"
```

Leave blank or comment out to use defaults. We won't judge... much.

### Output Directory

Tell it where to vomit your Excel files:
```env
PLEX_EXPORT_DIR="/path/to/export/directory"
```

Defaults to current directory if not specified. It's not hide-and-seek.

---

## 🤓 What You Need to Install (Dependencies, Not Issues)

```
plexapi>=4.15.4
pandas>=1.3.0
openpyxl>=3.0.9
requests>=2.26.0
python-dotenv>=0.19.0
```

Install all at once (because you're efficient like that):
```bash
pip install plexapi pandas openpyxl requests python-dotenv
```

---

## 🧠 Contributing (Join the Madness)

Want to help make this thing better? Great! Here's how to not screw it up:

1. Fork it (like a code buffet)
2. Branch it (`git checkout -b feature/MyAwesomeFeature`)
3. Commit it (`git commit -m 'Add something that doesn't break everything'`)
4. Push it (`git push origin feature/MyAwesomeFeature`)
5. Pull Request it (and pray I'm in a good mood)

Bonus points for witty commit messages.

---

## 🕊️ License (The Legal Mumbo Jumbo)

This thing is licensed under the **GNU AGPL v3.0**
It's free to use, but don't be a villain. [LICENSE](LICENSE)

---

## 👏 Credit Where It's Due (Props to the Real MVPs)

- [python-plexapi](https://github.com/pkkid/python-plexapi) - For making Plex integration not suck
- [TVMaze API](https://www.tvmaze.com/api) - For knowing more about TV than my mother
- [OpenPyXL](https://openpyxl.readthedocs.io/) - For making Excel files that don't crash
- [Pandas](https://pandas.pydata.org/) - For data processing that's smarter than me
- [Plex](https://www.plex.tv/) - For existing so we can hoard media legally-ish
- [nledenyi](https://github.com/nledenyi) - For contributing and not running away screaming
- [python-dotenv](https://github.com/theskumar/python-dotenv) - For managing environment variables better than I manage my life
- And you... yeah you, for actually reading this far

---

## 📨 Contact the Creator (Or Don't, I'm Not Your Boss)

Want to high-five the mastermind?

**PrimePoobah** — [GitHub](https://github.com/PrimePoobah)
**Project Link**: [PlexScripts](https://github.com/PrimePoobah/PlexScripts)

---

## 🔧 Troubleshooting & Logs

Something broke? Don't panic. We've got logs for that.

### Where Are The Logs?

The script creates detailed log files in the `logs/` directory:
- **File format:** `plex_export_YYYYMMDD.log`
- **Retention:** Automatically rotates at 5MB, keeps last 5 files
- **What's logged:** Everything - connections, API calls, errors, warnings, and debug info

### Checking The Logs
```bash
# View today's log (Linux/macOS)
cat logs/plex_export_$(date +%Y%m%d).log

# Watch logs in real-time (if running)
tail -f logs/plex_export_*.log

# Search for errors
grep "ERROR" logs/*.log
```

### Cache Files

The script creates a `.tvmaze_cache.pkl` file that stores TVMaze lookup results:
- **Location:** Export directory (same as your output files)
- **Lifetime:** 30 days (auto-expires old entries)
- **Size:** Usually a few KB to a few MB depending on your library
- **Delete it?** Sure! It'll rebuild on next run (just slower)

---

## ❓ FAQ (Frequently Awesome Questions)

### Q: Will this work on a headless server?
**A:** Yes, unlike me, this script doesn't need a pretty face to function.

### Q: Will this mess up my Plex library?
**A:** Nope, it's read-only. Less invasive than a wellness check.

### Q: How often should I run this?
**A:** Depends how often you add stuff. Weekly if you're obsessive, monthly if you have a life.

### Q: Can I customize the colors?
**A:** Sure, dive into the code and make it fabulous. Rainbow everything if you want. Edit the `STYLES` dictionary in [PlexMediaExport.py](PlexMediaExport.py). See [CLAUDE.md](CLAUDE.md) for architecture details.

### Q: Why don't some shows appear in TVMaze?
**A:** Because Plex and TVMaze sometimes disagree more than a married couple on vacation. The script shows `X/?` (yellow) for unknown data.

### Q: Where do the files go?
**A:** Same folder as the script, unless you specify otherwise in `.env`. It's not hide-and-seek.

### Q: Can I export only movies or only TV shows?
**A:** Currently, the script processes everything. You can manually delete unwanted worksheets from the Excel file, or dive into the code to filter sections. (Pull requests welcome!)

---

## ❤️ Show the Love

If this helped you organize your digital hoarding:

- ⭐️ Star this repo (make me feel special)
- Report issues on [GitHub Issues](https://github.com/PrimePoobah/PlexScripts/issues)
- Share it with other media addicts
- Contribute improvements (or just fix my terrible jokes)

---

Now go! Export that media like the spreadsheet superhero you were born to be.

*P.S. - If you're reading this far, you either really need this tool or you have way too much time on your hands. Either way, welcome to the club.*

*P.P.S. - No, I won't help you organize your actual physical media. That's what fire is for.*

**[END OF TRANSMISSION]**
