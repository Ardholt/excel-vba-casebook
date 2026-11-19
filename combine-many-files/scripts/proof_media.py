"""Replay measured invocation events; no simulated application controls or screen capture."""
import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

trace = json.loads(Path('evidence/refresh-trace.json').read_text())
report = json.loads(Path('output/report.json').read_text())
preview = json.loads(Path('evidence/table-preview.json').read_text())
events = trace['events']
finished = [e for e in events if e['phase'] == 'Complete']
assert len(finished) == 3
assert [e['rows'] for e in finished[:2]] == [report['output_rows']] * 2
font_dir = Path('/usr/share/fonts/truetype/dejavu')
def font(size, bold=False, mono=False):
    name = 'DejaVuSansMono' if mono else 'DejaVuSans'
    return ImageFont.truetype(str(font_dir / (name + ('-Bold' if bold else '') + '.ttf')), size)

W, H = 760, 714
ink, muted, accent = '#183F43', '#536C6F', '#B83D20'
def frame(event):
    im = Image.new('RGB', (W, H), '#F5F5F0')
    d = ImageDraw.Draw(im)
    def text(x, y, value, size=30, color=ink, bold=False, mono=False):
        f = font(size, bold, mono)
        assert d.textbbox((x, y), value, font=f)[2] <= W - 24, value
        d.text((x, y), value, font=f, fill=color)
    text(28, 20, 'REFRESH LOG  /  LIBREOFFICE', 25, bold=True)
    d.rounded_rectangle((24, 66, 736, 222), radius=12, fill=ink)
    names = {'first_refresh': 'First run', 'repeat_refresh': 'Run again',
             'added_file_refresh': 'Add a file, then run'}
    action = 'RUNNING' if event['phase'] != 'Complete' else 'COMPLETE'
    text(44, 81, names[event['operation']] + '  ·  ' + action, 25, '#C8E8DE', True)
    text(44, 119, '> RefreshFolder', 31, '#FFFFFF', mono=True)
    suffix = ' rows before refresh' if action == 'RUNNING' else ' rows in the table'
    text(44, 166, f"{event['rows']:,}" + suffix, 31, '#FFFFFF', True)
    text(28, 244, 'Recorded results', 26, bold=True)
    for i, (label, result) in enumerate(zip(['First run', 'Run again', 'Add 1 row + run'], finished)):
        y = 288 + 43*i
        text(28, y, label, 30)
        before = events[i*2]['rows']
        text(396, y, f"{before:,} → {result['rows']:,}", 30, bold=True)
    d.line((28, 426, 732, 426), fill='#CBD8D4', width=2)
    f = report['files']
    text(28, 443, f"{report['source_workbooks']} files: {f['accepted']} filled · {f['empty']} empty · {f['rejected']} rejected", 29, bold=True)
    text(28, 486, f"+ {f['ignored']} temporary file ignored (outside the {report['source_workbooks']})", 26, muted)
    d.line((28, 534, 732, 534), fill='#CBD8D4', width=2)
    row = preview['rows'][0]
    text(28, 550, 'One entry in the combined table', 26, bold=True)
    text(28, 591, f"{row['date']}   ·   Amount {row['amount']:.2f}", 30)
    text(28, 634, f"{row['source_file']}  ·  row {row['source_row']}", 27, mono=True)
    return im

Path('media').mkdir(exist_ok=True)
Path('images').mkdir(exist_ok=True)
frame(events[-1]).save('media/hero.png')
frames = [frame(event) for event in events]
# Retimed for reading: the raw trace retains the measured wall-clock timestamps.
durations = [1800 if e['phase'] != 'Complete' else 2400 for e in events]
durations[-1] = 4000
frames[0].save('images/refresh.gif', save_all=True, append_images=frames[1:],
               duration=durations, loop=0, disposal=2)
Path('evidence/proof-media.json').write_text(json.dumps({
    'kind': 'Animated replay of measured macro invocation events; not an application screen recording',
    'events': 'evidence/refresh-trace.json', 'counts': 'output/report.json',
    'sample': 'evidence/table-preview.json', 'frame_durations_ms': durations,
    'playback': 'Retimed for reading; measured elapsed times remain in the event trace.',
    'poster': 'media/hero.png', 'animation': 'images/refresh.gif',
    'dimensions': [W, H], 'font': 'DejaVu Sans / DejaVu Sans Mono',
}, indent=2) + '\n')
