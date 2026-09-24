# Ranking Nacional · Tenis de Mesa

A public website with the Costa Rica national table tennis rankings: every circuit (Liga Menor U9–U19, Liga Mayor, Open Femenino, Master, PTT), every stage, and each athlete's progress across stages and seasons.

## How data gets in

### 2026 onwards: automatic
Every Monday at 6:00 am (Costa Rica), GitHub checks the pages listed in `sources.json`, downloads any new .xlsx, converts it, and republishes the site. If FECOTEME publishes a stage and you do not want to wait until Monday, run it on demand: **Actions → Publish ranking → Run workflow**.

When the 2027 page appears, add one line to `sources.json`:
```json
{ "season": 2027, "url": "https://fecoteme.com/ranking-nacional-2027/" }
```

### 2025: one-time manual upload
The 2025 pages only offer downloads through buttons that need a browser, so these files have to be loaded by hand, once:

1. Download the files from each 2025 page:
   - Liga Menor: https://fecoteme.com/ranking-nacional-liga-menor-2025/
   - Liga Mayor: https://fecoteme.com/ranking-nacional-liga-mayor-2025/
   - Femenino: https://fecoteme.com/ranking-nacional-femenino-2025/
   - Master: https://fecoteme.com/ranking-nacional-master-2025/
   - PTT: https://fecoteme.com/ranking-nacional-ptt-2025/
2. On GitHub, open **`inbox/2025`**, choose **Add file → Upload files**, drag them all in, and click **Commit changes**. The folder name sets the season.
3. Check the run under **Actions**. Any file it couldn't classify stays in `inbox/2025` with an explanation in the run summary.

The same works for any extra file: drop it in `inbox/<year>/`.

**Keep the names the files download with.** The converter understands the 2025 naming (for example "Primer Ranking Nacional Master 60+", "3cer Ranking Nacional Femenino B", "4to Ranking Nacional Open Primera"). It also handles these quirks:
- **Open Femenino:** Primera = **Open A**, Segunda = **Open B** (the 2025 page uses both names).
- **PTT:** "3cer Ranking Nacional Master PTT" is filed under PTT, not Master.
- **Master:** a file with one sheet per bracket (30-39, 40-49, 50-59, 60+) is split into the four brackets.
- **"Primer Ranking Nacional Femenino":** the name doesn't say A or B. If the workbook doesn't say either, the run shows a warning. Add `Primer Ranking Nacional Femenino.xlsx.json` with `{"division": "Open A"}`, then upload the file again.

About 87 files in total for 2025: Menor 50, Mayor 5, Femenino 10, Master 17, PTT 5.

## How files are classified

| What | Where it comes from (most trusted first) |
|---|---|
| Category and gender | a `.json` override you add → the title inside the workbook (e.g. "U13 MASCULINO") → the FECOTEME link label → the file name |
| Stage | override → FECOTEME link label (`4to_ranking_…`) → file name (Roman numeral `IV-…`, or `Primer/Segundo/3er…`) |
| Season | override → `inbox/<year>/` folder → date in the workbook |
| Date | the date on the "Grupo" sheets |

When the FECOTEME label and the workbook (or file name) disagree, the run summary shows a **WARNING**, and warnings are kept in `data/warnings.json`. That has happened already: on the 2026 page, the link "1er ranking menor U9 masc" points to a file named "…U13-Femenino".

**Fixing a file:** add a file next to it with the same name plus `.json`, for example `inbox/2025/Primer-ranking-U13.xlsx.json`:
```json
{ "stage": 1, "division": "U13", "gender": "M", "circuit": "Liga Menor" }
```
Valid values:
- **division:** `U9`, `U11`, `U13`, `U15`, `U19`, `Mayor`, `Open A`, `Open B`, `Master`, `Master 40-49`, `PTT`
- **gender:** `M`, `F`, `X` (mixed)
- **circuit:** `Liga Menor`, `Liga Mayor`, `Open Femenino`, `Master`, `PTT`

**Re-uploading** the same season, circuit, category, gender and stage replaces the earlier version. **Removing** a stage means deleting its file in `data/stages/`.

## Athletes across years
Athletes are matched by **carné**, so their history carries over when they move from U11 to U13, or between circuits. Rankings from different categories have different sizes, so progress charts use the **percentile** (e.g. "Top 5%") rather than the raw position. Files without a carné column fall back to matching by name.

## Folders

| Folder | Contents |
|---|---|
| `inbox/<year>/` | Files waiting to be processed. Processed files move to `archive/<year>/`. |
| `data/stages/` | One JSON per stage (the source of truth) |
| `data/fetched.json` | FECOTEME links already downloaded |
| `site/` | The website; `site/data/ranking.json` is generated |
| `scripts/` | `fetch_fecoteme.py` (download), `convert.py` (convert) |

## One-time setup (GitHub Pages)
1. Create a **public** repository, e.g. `ranking-tm`, and upload this folder, including the hidden `.github` folder:
   ```bash
   cd ranking-tm && git init -b main && git add -A && git commit -m "Initial site"
   git remote add origin https://github.com/<user>/ranking-tm.git && git push -u origin main
   ```
2. Go to **Settings → Pages → Source** and choose **GitHub Actions**.
3. Go to **Actions → Publish ranking → Run workflow**. The first run downloads all of the 2026 files. The site is at `https://<user>.github.io/ranking-tm/`.
4. Optional: add a custom domain under **Settings → Pages**.

`site/` is a plain static site, so it can also be copied to S3 and CloudFront.
