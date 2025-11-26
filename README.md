# 📦 Plex Media Export Tools - Because Your Digital Media Hoard Deserves Excel-level OCD

![GPL License](https://img.shields.io/badge/license-GPL--3.0-blue)
![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![Plex API](https://img.shields.io/badge/PlexAPI-compatible-brightgreen)
![Excel Output](https://img.shields.io/badge/Excel-report-success)
![TVMaze Integration](https://img.shields.io/badge/TVMaze-integrated-informational)
![Pandas](https://img.shields.io/badge/Pandas-powered-ff69b4)
![Maintenance](https://img.shields.io/badge/maintained-yes-green.svg)
![Dotenv Support](https://img.shields.io/badge/dotenv-configurable-orange)
![Multi-Threading](https://img.shields.io/badge/processing-parallel-blue)

## Hey there, Beautiful Media Hoarders! 👋

So you've got a Plex server stuffed with more cinematic gold and TV trash, than a Blockbuster from 2003 (RIP, sweet prince), and you want to know what the hell you actually have? Well buckle up, buttercup! This isn't just another boring README written by some code monkey who thinks "fun" is a variable name.

This is a comprehensive collection of Python scripts that'll export your Plex content into Excel reports so detailed, they'd make the IRS weep tears of joy. We're talking movie catalogs, TV show completeness tracking, and enough color-coding to make a rainbow jealous.

---

## 🎞️ What Is This Sorcery?

Meet your new best frenemy: a trio of Python-powered bad boys that rip your Plex library a new spreadsheet.

- [**Plex Media Export**](https://github.com/PrimePoobah/PlexScripts/tree/main/Plex%20Media%20Export%20to%20Excel): The Swiss Army Chainsaw of Excel reports. Movies AND TV shows. Together. Harmony.
- [**Movie Exporter**](https://github.com/PrimePoobah/PlexScripts/tree/main/Plex%20Movie%20List%20Export%20to%20Excel): Just movies. Just the facts. Just justice.
- [**TV Show Audit Tool**](https://github.com/PrimePoobah/PlexScripts/tree/main/Plex%20TV%20Show%20Export%20to%20Excel): Completion stats, TVMaze brainpower, and a color-coded guilt trip for your binge-watching sins.

---

## 🧠 Features That Punch Mediocrity in the Face

### 🎬 For Movie Maniacs (Or: "How Many Times Did I Buy The Same Movie?")
- Complete inventory because you probably forgot you own Green Lantern (we don't talk about Green Lantern)
- Resolution-based highlighting that'll shame you for your 480p collection
- Technical details for the nerds (you know who you are)
- Content metadata so you can remember why you bought Cats (spoiler: there's no good reason)
- Customizable fields via .env because I'm not your mom – configure it yourself
- Alphabetical sorting because chaos is only fun in combat
- It's like IMDB got Excel-itis.

### 📺 TV Show Tracking (Or: "The 'Do I Actually Have All the Episodes?'")
- Finds missing episodes you didn't know you were avoiding.
- Series completion tracking with TVMaze integration (because manual counting is for psychopaths)
- Season-by-season breakdown that's more detailed than my therapy sessions
- Color-coded status indicators:
  - 🟩 Green: Complete (like my collection of regrets)
  - 🟥 Red: Incomplete (like my understanding of healthy relationships)
  - ⬛ Gray: This never existed (Like my social life)

### 💥 Advanced Features (Because We're Fancy Like That)
- **Multi-threaded processing** (faster than my mouth in a fight) - Parallel movie AND TV show processing
- **Persistent TVMaze cache** (90-95% faster on repeat runs) - Cached lookups survive between runs
- **Professional logging** (debugging without the headache) - Rotating log files with timestamps
- **Retry logic with exponential backoff** (network hiccups don't phase us)
- **Environment variable validation** (catches config errors before you waste time)
- **Environment variables from .env files** (because hardcoding is what killed the dinosaurs)
- **Memory-optimized Excel generation** (won't crash your potato computer)
- **Progress reporting** more detailed than my therapy notes
- **Sortable Excel tables** (organization is my middle name... actually it's Winston, but whatever)

---

## 🚀 Installation Instructions (or “How to Not Screw It Up”)

### 🧰 Requirements
- Python 3.8+ (we're modern now, grandpa)
- Plex server (not imaginary)
- Plex token (like a golden key but nerdier)
- Internet (for TVMaze, not for my OnlyFans)

### 🧙‍♂️ Summon the Tools

```bash
git clone https://github.com/PrimePoobah/plex-media-export.git
cd plex-media-export
pip install plexapi pandas openpyxl requests python-dotenv
cp .env.example .env
nano .env  # Or edit in Notepad if you're stuck in 2005
```

**Configure your settings (this is where the magic happens):**
```
PLEX_URL="http://{Your_Plex_IP}:32400"
PLEX_TOKEN="{YourSuperSecretPlexToken}"
PLEX_EXPORT_DIR="{WhereYouWantYourFiles}"
PLEX_MOVIE_EXPORT_FIELDS="Title,Year,Studio,ContentRating,Video Resolution,File Path,Container,Duration (min)"
PLEX_SHOW_EXPORT_FIELDS="Title,Year,Studio,ContentRating"
```

### 🎟️ PLEX_TOKEN Hunt
1. Log into Plex web interface (you know, that thing you use to procrastinate)
2. Play any media file (I suggest *Deadpool*)
3. Click the three dots (⋮) because apparently we're all ancient Greeks now
4. Select "Get Info" (getting philosophical, are we?)
5. Click "View XML" (now we're speaking in tongues)
6. Find "X-Plex-Token" in the URL (congratulations, you're now a hacker)

---

## 🎯 Running the Scripts

### 1️⃣ The Big Kahuna

This script does everything. It's like the Swiss Army knife of Plex tools, if Swiss Army knives could judge your media collection.

```bash
python PlexMediaExport.py
```
Output: `PlexMediaExport_YYYYMMDD_HHMMSS.xlsx`
*Because timestamps are like signatures – they prove you were here*
Contains both:
- **Movies** tab
- **TV Shows** tab (with Maze magic)

### 2️⃣ Just the Movies, Ma’am

For when you only care about movies. It's focused, dedicated, and probably has commitment issues.

```bash
python plex_movie_export.py
```
Output: `plex_movies.xlsx`
*Simple name for simple people*

### 3️⃣ Show Me the TV

This one tracks TV show completion like a particularly obsessive stalker.

```bash
python plex_tv_shows.py
```
Output: `plex_tv_shows_YYYYMMDD.xlsx`

*With a timestamp because even TV shows deserve to know when they were cataloged*

---

## 📊 Spreadsheet Glory Details

### Movies Sheet

| Column       | Description         |
|--------------|---------------------|
| Title        | Movie name          |
| Resolution   | 4K? Fancy. SD? Eww. |
| Year         | Release year        |
| Studio       | Studio magic        |
| File         | Full path to shame  |
| Container    | MKV? MP4? VHS?      |
| Duration     | Minutes of regret   |
| Etc.         | You decide via `.env` |

### TV Shows Sheet

| Column        | Description                 |
|---------------|-----------------------------|
| Title         | Show name                   |
| Complete      | Yes, no, or… yikes          |
| Season X      | Episodes present/total      |

---

## 🎨 Color Legend (Read 'em and Weep)

### Movies
- 🟩 4K/UHD = 4K content (oh shiny!)
- 🟨 720p or lower (we don't judge... much)
- ⬜ 1080p (the standard bearer of mediocrity)

### Shows
- 🟩 Season done. Victory lap.
- 🟥 You slacked off.
- ⬛ Didn’t even exist. Move on.

---

## 🔧 Customization (Make It Your Own, Like A Bad Tattoo)

Edit `.env` like a wizard.

### 🎥 Movie Fields (Choose your destiny)
```
Title, Year, Studio, Rating, Bitrate, Codec, Genres, Labels...
```

### 📺 TV Show Fields (Choose your own binge-venture)
```
Title, Seasons, Ratings, ViewCount, Summary, Tagline...
```

---

## 🤓 What You Need to Install (Dependencies, Not Issues)

```
plexapi>=4.15.4
pandas>=1.3.0
openpyxl>=3.0.9
requests>=2.26.0
python-dotenv>=0.19.0
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

## 🕊️  License (The Legal Mumbo Jumbo)

This thing is licensed under the **GNU AGPL v3.0**  
It’s free to use, but don’t be a villain. [LICENSE](LICENSE)

---

## 👏 Credit Where It’s Due (Props to the Real MVPs)


- [python-plexapi](https://github.com/pkkid/python-plexapi) For making Plex integration not suck
- [TVMaze API](https://www.tvmaze.com/api) For knowing more about TV than my mother
- [OpenPyXL](https://openpyxl.readthedocs.io/) For making Excel files that don't crash
- [Pandas](https://pandas.pydata.org/) For data processing that's smarter than me
- [Plex](https://www.plex.tv/) For existing so we can hoard media legally-ish
- [nledenyi](https://github.com/nledenyi) For contributing and not running away screaming
- [python-dotenv](https://github.com/theskumar/python-dotenv) For managing environment variables better than I manage my life
- And you... yeah you, for actually reading this far

---

## 📨 Contact the Creator (Or Don't, I'm Not Your Boss)

Want to high-five the mastermind?

**PrimePoobah** — [GitHub](https://github.com/PrimePoobah)  
**Project Link**: [plex-media-export](https://github.com/PrimePoobah/plex-media-export)

---

## 🔧 Troubleshooting & Logs

Something broke? Don't panic. We've got logs for that.

### Where Are The Logs?
The script creates detailed log files in the `logs/` directory (or wherever you set `PLEX_EXPORT_DIR`):
- **File format:** `plex_export_YYYYMMDD.log`
- **Retention:** Automatically rotates at 5MB, keeps last 5 files
- **What's logged:** Everything - connections, API calls, errors, warnings, and debug info

### Checking The Logs
```bash
# View today's log
cat logs/plex_export_$(date +%Y%m%d).log

# Watch logs in real-time (if running)
tail -f logs/plex_export_$(date +%Y%m%d).log

# Search for errors
grep "ERROR" logs/*.log
```

### Cache Files
The script creates a `.tvmaze_cache.pkl` file that stores TVMaze lookup results:
- **Location:** Same as your export directory
- **Lifetime:** 30 days (auto-expires old entries)
- **Size:** Usually a few KB to a few MB depending on your library
- **Delete it?** Sure! It'll rebuild on next run (just slower)

---

## ❓FAQ (Freakin' Awesome Questions)

### Q: Will this work on a headless server?
**A:** Yes, unlike me, these scripts don't need a pretty face to function.

### Q: Will this mess up my Plex library?
**A:** Nope, it's read-only. Less invasive than a wellness check.

### Q: How often should I run these?
**A:** Depends how often you add stuff. Weekly if you're obsessive, monthly if you have a life.

### Q: Can I customize the colors?
**A:** Sure, dive into the code and make it fabulous. Rainbow everything if you want.

### Q: Why don't some shows appear?
**A:** Because Plex and TVMaze sometimes disagree more than a married couple on vacation.

### Q: Where do the files go?
**A:** Same folder as the script, unless you specify otherwise. It's not hide-and-seek.

---

## ❤️ Show the Love

If this helped you organize your digital hoarding:

- ⭐️ Star this repo (make me feel special)
- Share it with other media addicts
- Contribute improvements (or just fix my terrible jokes)

  ---

Now go! Export that media like the spreadsheet superhero you were born to be.

*P.S. - If you're reading this far, you either really need this tool or you have way too much time on your hands. Either way, welcome to the club.*

*P.P.S. - No, I won't help you organize your actual physical media. That's what fire is for.*

**[END OF TRANSMISSION]**

*This README was written while consuming an ungodly amount of chimichangas and watching my own movies on repeat. Any resemblance to actual documentation is purely coincidental.*
