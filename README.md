# substack-stacker

Turn any Substack into a single .docx file. All posts, one document, like a little book.

## Setup

```bash
pip install -e .
```

## Usage

```bash
# just pass a substack url or name
python stacker.py https://example.substack.com

# or just the name
python stacker.py example

# chronological order (oldest first)
python stacker.py example --oldest-first

# only grab the last 10 posts
python stacker.py example --limit 10

# custom output filename
python stacker.py example -o my_book.docx
```

The output is a .docx with a title page, table of contents, and each post as its own chapter with a page break in between. Images, formatting, code blocks, blockquotes are all preserved!

## Options

| Flag | What it does |
|------|-------------|
| `-o FILE` | Output filename (default: `{name}_posts.docx`) |
| `--oldest-first` | Chronological order instead of newest first |
| `-n N` / `--limit N` | Only fetch N posts |
| `--delay SECS` | Delay between API requests (default: 1.5s) |

## Examples

Check the `examples/` folder for a sample output.
