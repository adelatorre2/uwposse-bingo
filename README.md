# Posse Bingo

A digital, phone-friendly version of Posse Bingo for UW–Madison Posse SOAR. No app, no printing — scholars open a link (or scan a QR code), tap squares as they're true for someone in their Posse, and race to five in a row.

Live at: https://adelatorre2.github.io/uwposse-bingo/

## How it works

- Two decks, switchable at the top: **Warm Up** (lighthearted icebreaker squares — siblings, pets, hobbies) and **Go Deeper** (identity, first-gen, PWI, and belonging-focused squares). This mirrors the "Warm up / Go deeper" language from the [dyad questions site](https://adelatorre2.github.io/uwposse-dyad-questions/).
- Every phone gets its **own randomly shuffled board** per deck (persisted in that browser via `localStorage`), so no two scholars are staring at an identical grid and a "bingo" actually means something.
- The center square is a fixed "connector" square per deck (e.g. *Shares something in common with you*) — it still has to be tapped, it's just always in the middle to match the printed cards.
- Tapping a square toggles it marked/unmarked. Getting five in a row (row, column, or diagonal) triggers a confetti celebration; a full card triggers a "Blackout" celebration.
- **New board** reshuffles that deck's squares into a fresh grid. **Clear marks** resets the taps without reshuffling.

## Editing the squares

No build step — just edit `data/bingo_items.xlsx` (sheet name **Items**, columns `deck`, `text`, `is_center`), commit, and push. The site re-reads the file at runtime.

- `deck` — which deck the square belongs to (`Warm Up` or `Go Deeper`, or a new deck name — new decks automatically get a toggle button if you add markup for it, otherwise add a button in `index.html`).
- `text` — the square's text.
- `is_center` — `TRUE` for exactly one row per deck (the square that always lands in the middle); leave blank/`FALSE` for the rest.

## Local preview

Any static file server works, e.g.:

```
python3 -m http.server 8000
```

then open `http://localhost:8000`.

## Credits

Bingo squares originally designed in Canva for Posse SOAR. Built by [Alejandro De La Torre](https://adelatorre2.github.io) — an independent project; views expressed do not represent the UW–Madison Posse Program.
