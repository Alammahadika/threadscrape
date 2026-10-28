# ThreadScrape
### A Python-Based Threads Search Data Collector

**ThreadScrape** is a Python-based data collection tool designed to collect publicly displayed posts from Threads search results using keyword-based queries. Built with Playwright and Beautiful Soup, it extracts post metadata and exports the collected data into structured Excel (`.xlsx`) files.

The tool is designed to support social media research, computational social science, social computing, digital discourse analysis, and other research projects that require structured data from Threads.

> **Project status:** Under development. The scraper is functional, but its reliability, data completeness, and compatibility with changes to the Threads interface are still being evaluated.

## Features

- **Keyword-based search:** Collect posts from Threads search results using user-defined keywords or phrases.
- **Automated scrolling:** Navigate search results and progressively collect additional posts.
- **Post identification:** Extract post IDs, usernames, and post URLs.
- **Text extraction:** Collect the text displayed in individual posts.
- **Metadata extraction:** Retrieve available timestamps, verification indicators, topic tags, and engagement metrics.
- **Media detection:** Identify posts containing detectable images or videos.
- **Duplicate handling:** Identify duplicate posts using their post IDs.
- **Excel export:** Save collected data in structured `.xlsx` files.
- **Offline HTML parsing:** Parse saved HTML files without launching a live browser session.
- **Session persistence:** Save an authenticated browser session for subsequent scraping runs.

## Technology Stack

ThreadScrape uses the following Python libraries:

| Technology | Purpose |
|---|---|
| Python | Core programming language |
| Playwright | Browser automation and live data collection |
| Beautiful Soup 4 | HTML parsing and post extraction |
| lxml | HTML parsing backend |
| pandas | Data organization and processing |
| openpyxl | Excel file generation and formatting |

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/Alammahadika/threadscrape.git
cd threadscrape
```

### 2. Install dependencies

```bash
pip install beautifulsoup4 lxml pandas openpyxl playwright
```

### 3. Install the Playwright browser

```bash
playwright install chromium
```

## Authentication

Live scraping requires a saved browser session.

Run:

```bash
python threadscrape.py --login
```

A Chromium browser window will open. Log in to Threads manually, then return to the terminal and press Enter to save the browser session.

The session is saved in:

```text
threads_state.json
```

Subsequent scraping runs use this saved session.

**Security note:** Treat `threads_state.json` as sensitive authentication data. Do not publish it, commit it to a public repository, or share it with others. Add it to `.gitignore`.

## Usage

ThreadScrape supports two collection modes: live browser scraping and offline HTML parsing.

### 1. Live scraping

Collect posts directly from Threads search results using a keyword query.

```bash
python threadscrape.py --query "forest fire" --scrolls 100
```

Parameters:

| Parameter | Description |
|---|---|
| `--query` | Search keywords or phrases |
| `--scrolls` | Maximum number of scrolling iterations |
| `--headed` | Display the browser window during scraping |
| `--out` | Specify the output Excel filename |
| `--patience` | Stop after a specified number of consecutive iterations without newly collected posts |
| `--login` | Authenticate and save the browser session |

Example with a visible browser:

```bash
python threadscrape.py --query "prabowo subianto" --scrolls 400 --headed
```

Example with a custom output filename:

```bash
python threadscrape.py \
  --query "forest fire" \
  --scrolls 100 \
  --out forest_fire_data.xlsx
```

Example with a custom stopping threshold:

```bash
python threadscrape.py \
  --query "forest fire" \
  --scrolls 400 \
  --patience 8
```

The `--scrolls` parameter specifies the maximum number of scrolling iterations, not the number of posts to collect. The number of collected posts depends on search results, page behavior, and the number of unique posts successfully extracted.

### 2. Offline HTML parsing

ThreadScrape can also parse previously saved HTML files.

```bash
python threadscrape.py --html results.html --query "forest fire"
```

Multiple HTML files can be processed in a single run:

```bash
python threadscrape.py \
  --html results_1.html results_2.html results_3.html \
  --query "forest fire"
```

This mode is useful for processing saved search-result pages without performing live browser automation.

The HTML files must contain the relevant post elements and metadata required by the parser.

## Data Collection Workflow

The collection process follows these general steps:

1. **Search initialization:** Open the Threads search page using the specified query.
2. **Initial page loading:** Wait for the search-result elements to appear.
3. **HTML extraction:** Read the currently rendered page and parse available post elements.
4. **Post identification:** Extract post IDs and identify previously collected posts.
5. **Duplicate handling:** Add only posts whose IDs have not already been collected during the current run.
6. **Scrolling:** Move through the search results to trigger additional content loading.
7. **Progress monitoring:** Report newly collected posts and the cumulative collection count.
8. **Stopping condition:** Stop when the maximum scroll count is reached or the configured consecutive no-new-post threshold is met.
9. **Data organization:** Convert the collected records into a structured pandas DataFrame.
10. **Excel export:** Save the resulting dataset and collection metadata in an Excel workbook.

## Output Data

ThreadScrape exports the collected data into an Excel workbook containing two worksheets.

### Worksheet 1: `posts`

| Column | Description |
|---|---|
| `query` | Search query used during collection |
| `post_id` | Unique identifier extracted from the post URL |
| `username` | Account username |
| `verified` | Whether a detectable verification indicator was found |
| `timestamp` | Post timestamp converted from the extracted timestamp to UTC |
| `date` | Date derived from the timestamp |
| `text` | Extracted post text |
| `hashtags` | Hashtags extracted from the post text |
| `mentions` | Account mentions extracted from the post text |
| `topic_tag` | Topic tag when available |
| `media` | Detected media type, such as image or video |
| `likes` | Extracted like count when available |
| `replies` | Extracted reply or comment count when available |
| `reposts` | Extracted repost count when available |
| `shares` | Extracted share count when available |
| `engagement` | Sum of the extracted likes, replies, reposts, and shares |
| `url` | Direct URL of the post |

The `text` field retains the extracted text with only minimal whitespace normalization. Content cleaning, tokenization, sentiment analysis, and other analytical preprocessing are intentionally left to downstream workflows.

**Data interpretation note:** Engagement metrics may be unavailable or incorrectly extracted if Threads changes its interface or the relevant labels. A value of zero does not necessarily mean that the post received no engagement.

### Worksheet 2: `info`

The `info` worksheet records collection metadata, including:

- Search query
- Collection mode
- Number of posts in the exported dataset
- Timestamp range in UTC
- Collection date and time

These fields help document the collection process and support research reproducibility.

## Example Output

A typical collection run produces terminal logs similar to the following:

```text
scroll 1/400: +5 (total 5, stale 0/8)
scroll 2/400: +3 (total 8, stale 0/8)
scroll 3/400: +4 (total 12, stale 0/8)
...
scroll 400/400: +2 (total 1779, stale 0/8)

Stopped: reached 400 scrolls

OK: 1772 posts -> threads_kopdes_20261027_2019.xlsx
```

In this example, the scraper discovered 1,779 unique post IDs during collection, while the final dataset contained 1,772 posts after processing.

The difference may result from records being removed during data preparation, including posts without detectable text or media. The collection count and final export count should therefore be treated as separate measurements.

The number of scrolling iterations does not guarantee that all available search results have been collected.

## Intended Applications

ThreadScrape can support research and analytical workflows involving:

- Social computing and computational social science
- Public discourse and political communication
- Social media content analysis
- Digital ethnography and online discourse research
- Public sentiment and opinion analysis
- Topic modeling and text mining
- Social media research data preparation

The exported dataset can be used with analytical tools such as Python, R, NVivo, and other software that supports structured data.

## Limitations

ThreadScrape is an experimental research-oriented collection tool and has several limitations.

**1. Dynamic interface**

Threads uses a dynamic web interface. Changes to HTML elements, selectors, rendering behavior, or page structure may affect data extraction.

**2. Search-result coverage**

The scraper collects posts that it can access and extract from the rendered search results. It does not guarantee complete coverage of all posts matching a query.

**3. Scrolling behavior**

A scrolling iteration is not equivalent to a post, a page of fixed size, or a guaranteed quantity of newly collected records.

**4. Authentication requirements**

Live collection requires an authenticated browser session. Changes to login requirements or session validity may interrupt execution.

**5. Metadata availability**

Some post fields and engagement metrics may be missing, inaccessible, or incorrectly extracted.

**6. Duplicate handling**

Duplicate detection is based on `post_id` within a single collection run. Separate runs may contain overlapping records and require additional deduplication.

**7. Research validity**

A large number of collected posts does not automatically establish dataset completeness, representativeness, or validity. Researchers should document collection parameters, evaluate data quality, and assess the limitations of the sampling process.

## Responsible Use

Use ThreadScrape in accordance with applicable laws, Threads' terms and policies, and relevant research ethics requirements.

Collect only data that you are authorized to access. Avoid excessive requests, respect applicable access restrictions, and protect authentication credentials and sensitive information.

Researchers should consider user privacy, data minimization, informed ethical review where applicable, and the potential risks of publishing identifiable social media content.

## Project Status

ThreadScrape is under active development.

Current development priorities include improving the reliability of browser interactions, handling dynamic page changes, monitoring collection progress, and improving the reproducibility of data collection.

Future improvements may include:

- Periodic checkpointing and recovery after interruptions
- More robust handling of dynamic DOM changes
- Improved collection diagnostics and logging
- Stronger validation of extracted post metadata
- Collection summaries and data-quality reports
- Additional export and downstream processing options

## License

A license has not yet been specified in this README. Add an appropriate `LICENSE` file to the repository before distributing the project under an open-source license.

## Author

**Alam Mahadika**

GitHub: https://github.com/Alammahadika

---

*ThreadScrape is intended to support structured data collection for research and analytical purposes. It is not an official Threads product and is not affiliated with or endorsed by Meta.*
