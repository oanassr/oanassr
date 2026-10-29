# Voice cloning steps (run on your own machine)

1. Install Voice Studio from https://aiaudiogen.com and `pip install uv`.
2. Open this repo locally in Claude Code; approve the `voicestudio` server from `.mcp.json`.
3. Collect clean samples of YOUR OWN voice (YouTube is blocked in the cloud session):
   - yt-dlp -x --audio-format wav <url>   (for the three videos you own)
   - Keep 1-3 minutes of clean speech, no music or overlap, trim in ffmpeg:
     ffmpeg -i in.wav -ar 24000 -ac 1 -ss 00:00:10 -t 120 sample.wav
4. Ask Claude: "Clone my voice from sample.wav with voicestudio, then narrate script.md".
