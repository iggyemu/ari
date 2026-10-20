# Ari’s 2026 holiday wishlist

The shareable page is **https://iggyemu.github.io/ari/wishlist.html**.
It is a static HTML page: no application server, build step, or JavaScript is required for visitors.

## Publish to your existing GitHub Pages site

This working folder does not contain Git metadata. Use your existing clone of `iggyemu/ari`,
or clone that repository into a separate folder using GitHub Desktop or Git. Keep the repository’s
current Pages configuration and publishing branch.

Copy these files into the folder that currently publishes `workout.html`, preserving their paths:

```text
wishlist.html
assets/wishlist-2026/video.mp4
assets/wishlist-2026/poster.jpg
dependencies/wishlist 2026 ari.pptx
```

Also keep `scripts/build-wishlist.py`, `scripts/export-wishlist.ps1`, and this document in the repository
for future updates. Add `.wishlist-build/` and `__pycache__/` to its existing `.gitignore`.
Do not upload `.wishlist-build/`: it contains temporary presentations, native exports, QA images,
and downloaded tools. The original `workout.html` does not need modification.

Commit and push with Git or GitHub Desktop. The original presentation is about 36 MiB,
which exceeds GitHub’s 25 MiB browser-upload limit. Keep the MP4 below 95 MiB and use ordinary
Git storage; Git LFS is not supported for GitHub Pages.

Once the existing Pages deployment finishes:

1. Open the published wishlist URL on a computer and a phone.
2. Play, pause, seek near the end, and enter fullscreen. Confirm the thank-you message appears.
3. Download the PowerPoint and open it in PowerPoint.
4. Try the direct MP4 link and several shopping links.
5. Confirm the existing workout page still loads.

If an asset returns 404, check that it was copied to the same publishing root as `wishlist.html`
and that its filename, letter case, and directory match. No leading `/` is used for page assets,
so the `/ari/` project path is retained.

## Regenerate the gift list

Requires Python 3 with no additional packages:

```powershell
python .\scripts\build-wishlist.py
```

The script reads the original PPTX and updates only the marked gift-list section in `wishlist.html`.
It handles visible URLs, embedded hyperlinks, split text runs, and the multi-product jewelry and
golf slides. It unwraps only the explicit destinations in the deck’s advertising redirects.
It does not search for replacement products or invent missing links.
Sizes appear beneath the product names where supplied, preserving pant length.
On the webpage, “xss” and “xxs” display as “XXS”, and “Women’s 5.5” and “Women 5.5”
display as “Women's 5.5”. These display corrections do not change the original PowerPoint.
The Ashlyn striped top (slide 4) displays the family-supplied size XXS on the webpage.
Family-supplied colors also appear on the webpage: Light Blue for the TaylorMade golf bag,
and Summit White/Court Blue/Metallic Silver/Obsidian for the Nike Victory Pro 4 shoes.

There are 33 gift entries: 15 Clothes, 10 Items, and 8 Golf. Eighteen have a source product link;
the remaining fifteen show the retailer and “Product link not provided.” Product wording and
destinations come from the source, including the two shorts slides that share the same title
but have different product URLs. Availability, prices, and whether each destination still sells
the intended product are not independently verified.

If slides are added, removed, or reorganized, review the slide ranges and special table/shape
mappings in the script before regenerating. Its source-count check stops if the count changes.

## Regenerate the video

Requires Windows, installed desktop Microsoft PowerPoint, and FFmpeg plus FFprobe.
Download a portable Windows build through [FFmpeg’s download page](https://ffmpeg.org/download.html);
the verified build used here is Gyan’s FFmpeg 9.0.2 essentials build.
Extract it under `.wishlist-build/tools/`, add its `bin` folder to PATH, or supply both executable paths.
The script does not download or install software itself.

```powershell
.\scripts\export-wishlist.ps1 `
  -FfmpegPath 'C:\path\to\ffmpeg.exe' `
  -FfprobePath 'C:\path\to\ffprobe.exe'
```

When the tools are already in PATH or under `.wishlist-build/`, simply run:

```powershell
.\scripts\export-wishlist.ps1
```

PowerPoint operates on `.wishlist-build/wishlist-export-copy.pptx`. The original is never saved,
and its SHA-256 is checked after export. Slides 1–34 have a five-second reading interval, with
their existing transitions retained. The closing sequence plays the golf clip once, reveals
and plays the wave clip once in its original position, then fades in the existing thank-you text.
The wave is brought above the golf background so it can be seen. Existing mute settings are preserved;
this deck produces a silent video. Only the export copy loses the slide-34 click instruction.
PowerPoint’s automatic date fields display the export date.

The native export is 1920 × 1080 at 30 fps. FFmpeg encodes H.264, `yuv420p`, and fast-start MP4.
If a quality-based encode exceeds 95 MiB, the script uses a two-pass encode targeting 90 MiB.
Publication is blocked if the dimensions, codec, frame rate, pixel format, or size do not match.
Existing final assets are replaced only after the video passes those checks.

Useful options:

```powershell
# Export just the closing sequence to the temporary folder for inspection.
.\scripts\export-wishlist.ps1 -EndingPreview

# Recompress an existing full native export without reopening PowerPoint.
.\scripts\export-wishlist.ps1 -SkipPowerPoint
```

Do not run exports concurrently: they share the working-copy filename.
After changing the source presentation, run both scripts and review the video and gift links.
The page styling can be edited directly; the gift-list generator preserves everything outside its markers.

## Validation

The delivered MP4 is **4:14.23**, **26.0 MiB**, H.264 / `yuv420p`, **1920 × 1080 at 30 fps**,
with fast-start metadata and no audio track. Every frame decoded successfully, and video samples
cover all 35 slides in order plus the automatic closing sequence.

Local browser checks covered playback, keyboard seeking, 1280-pixel desktop and 390-pixel phone
layouts, category navigation, and the PowerPoint download. The downloaded file’s SHA-256 matches
the source, and the source presentation and workout page are unchanged. A simulated missing-video
response with scripts blocked left the downloads and all 33 gift entries available.
Responsive checks used viewport emulation; real-phone and live Pages checks remain after deployment.

Inspect all slides and the closing sequence after each source change, especially fonts,
long product descriptions, slide order, and transitions. Use fullscreen for the small text
already present in the source deck. Native PowerPoint rendering preserves the deck’s layout.

Check the compressed file using FFprobe:

```powershell
ffprobe -v error -show_format -show_streams -of json .\assets\wishlist-2026\video.mp4
```

Test the page at desktop and narrow mobile widths with keyboard navigation and JavaScript disabled.
The native player has no autoplay, uses inline playback on phones, and preloads only metadata.
The direct MP4 and original PowerPoint links remain visible if embedded playback fails.

The page is published through the repository’s existing Pages pipeline. Repeat the live checks
above after future updates. The video links use a version query to refresh cached playback files.
