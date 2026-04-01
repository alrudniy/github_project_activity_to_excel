# GitHub Project V2 → Excel Exporter

Export all items from a **GitHub Project V2** board into a formatted Excel spreadsheet. Captures issues, pull requests, and draft issues — including body, comments, code reviews, linked PRs, custom project fields, and more.

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Authentication](#authentication)
- [Usage](#usage)
- [Output Format](#output-format)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)

---

## Features

- Exports all item types: **Issues**, **Pull Requests**, and **Draft Issues**
- Handles **pagination** automatically (fetches all items regardless of project size)
- Extracts **19 fixed columns** plus any **custom project fields** you've defined
- Captures rich detail per item:
  - Author, editor, assignees, labels, milestone
  - Full body text, comments (up to 50), code reviews and review comments (up to 50)
  - Linked PRs and closing references
  - Created/updated timestamps and item URL
- Supports both **organization-owned** and **user-owned** projects
- Produces a professionally styled Excel file:
  - Dark blue header row with white bold text
  - Frozen header row and auto-filters on all columns
  - Auto-fitted column widths; fixed widths for long-text columns
  - Wrapped text in body, comments, and review columns
  - Alternating gray row shading for readability

---

## Requirements

- Python 3.8+
- A GitHub Personal Access Token (PAT) with **`read:project`** scope

### Python packages

```
requests
openpyxl
```

---

## Installation

1. **Clone or download** this repository:

   ```bash
   git clone <repo-url>
   cd "!2026-03-22  Extract github project details"
   ```

2. **Create and activate a virtual environment:**

   ```bash
   python -m venv venv

   # On Windows (Git Bash):
   source venv/Scripts/activate

   # On macOS/Linux:
   source venv/bin/activate
   ```

3. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

---

## Authentication

The script requires a GitHub Personal Access Token (PAT) set as an environment variable.

### Creating a PAT

1. Go to **GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens** (or classic tokens).
2. Generate a new token with the **`read:project`** scope (and `repo` if you need issue/PR bodies from private repos).
3. Copy the token value — you won't be able to see it again.

### Setting the token

```bash
export GITHUB_TOKEN="ghp_yourtoken"
```

Add this to your shell profile (`.bashrc`, `.zshrc`, etc.) to avoid setting it every session.

> The script will exit with an error message if `GITHUB_TOKEN` is not set.

---

## Usage

```
python github_project_to_excel.py --owner <owner> --project <number> [--output <file.xlsx>] [--user]
```

### Arguments

| Argument | Required | Default | Description |
|---|---|---|---|
| `--owner` | Yes | — | GitHub organization name or username that owns the project |
| `--project` | Yes | — | Project number (the integer in the project URL: `github.com/orgs/<org>/projects/<number>`) |
| `--output` | No | `github_project_items.xlsx` | Path/filename for the output Excel file |
| `--user` | No | *(flag, off by default)* | Pass this flag when the project is owned by a **user** instead of an **organization** |

### Quick start

```bash
# Activate venv and set token
source venv/Scripts/activate
export GITHUB_TOKEN="ghp_yourtoken"

# Export project #6 from organization "my-org"
python github_project_to_excel.py --owner my-org --project 6

# Export with a custom output filename
python github_project_to_excel.py --owner my-org --project 6 --output my_project_export.xlsx

# Export a project owned by a user (not an org)
python github_project_to_excel.py --owner myusername --project 3 --user
```

---

## Output Format

The generated `.xlsx` file contains one row per project item with the following columns:

### Fixed columns (always present)

| Column | Description |
|---|---|
| Author | GitHub username of the item creator |
| Editor | GitHub username who last edited the body |
| Type | `Issue`, `PullRequest`, or `DraftIssue` |
| Title | Item title |
| # | Issue/PR number (blank for draft issues) |
| State | `OPEN`, `CLOSED`, `MERGED`, etc. |
| Repository | `owner/repo` for issues and PRs |
| Assignees | Comma-separated list of assignees |
| Labels | Comma-separated list of labels |
| Milestone | Milestone title if assigned |
| Body | Full body/description text |
| Comments Count | Total number of comments |
| Comments | All comments formatted as `[author date]: text`, separated by `---` |
| Reviews | PR code reviews with state (APPROVED / CHANGES_REQUESTED / COMMENTED) and inline review comments |
| Linked PRs/Issues | Cross-referenced PRs and issues this item closes |
| Created | ISO timestamp of creation |
| Updated | ISO timestamp of last update |
| URL | Direct link to the item on GitHub |

### Dynamic columns

Any **custom fields** configured on your GitHub Project V2 board (text, number, date, single-select, iteration) are appended as additional columns after the fixed ones.

---

## Examples

```bash
# Export an opportunity-tracking project
python github_project_to_excel.py \
  --owner my-org \
  --project 6 \
  --output project_opportunity_$(date +%Y%m%d).xlsx

# Export a personal user project
python github_project_to_excel.py \
  --owner myusername \
  --project 2 \
  --user \
  --output personal_project_export.xlsx
```

---

## Troubleshooting

**`Error: GITHUB_TOKEN environment variable not set.`**
Set your token: `export GITHUB_TOKEN="ghp_yourtoken"`

**`HTTP error: 401`**
Your token is invalid or expired. Generate a new PAT in GitHub settings.

**`HTTP error: 403`**
Your token lacks the required `read:project` scope. Regenerate with the correct permissions.

**`GraphQL error: Could not resolve to a ProjectV2...`**
- Double-check `--owner` and `--project` match the GitHub URL exactly.
- If the project belongs to a user (not an org), add the `--user` flag.
- Ensure your PAT has access to the organization if the project is in a private org.

**Output file is empty / missing rows**
The script fetches up to 100 items per page and paginates automatically. If the project is very large (thousands of items), the fetch may take a minute. Check the terminal output for progress.

**Custom fields not appearing**
Custom fields are fetched dynamically. If a field is empty for all items, it will still appear as a column header but with blank values.
