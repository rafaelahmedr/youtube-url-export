# YouTube Channel URL Exporter

Fetch all video titles and URLs from a YouTube channel and save them into an Excel file.

## Features
- Extracts all video titles from any public YouTube channel  
- Saves data into Excel (`.xlsx`) with clickable links  
- Works with any public channel (no API key required)  
- Simple and lightweight, uses [yt-dlp](https://github.com/yt-dlp/yt-dlp) under the hood  

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/rafaelahmedr/youtube-channel-export.git
   cd youtube-channel-export
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the script with a channel URL:

```bash
python yt_to_excel.py https://www.youtube.com/@CotesAS
```

The output will be saved as: **videos.xlsx**

## Example Output

| Title            | URL                                     |
|------------------|-----------------------------------------|
| Example Video 1  | https://www.youtube.com/watch?v=abc123  |
| Example Video 2  | https://www.youtube.com/watch?v=def456  |

## Notes
- The script works only with public videos  
- For members-only, unlisted, or private content, authentication is required (not included in this version)  
- Update yt-dlp regularly to ensure compatibility with YouTube changes:
  ```bash
  pip install -U yt-dlp
  ```

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
