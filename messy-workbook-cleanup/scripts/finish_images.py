"""Assemble the two Calc sheet renders without adding or changing evidence."""
import json
import shutil
from pathlib import Path
from PIL import Image


def main():
    pictures = []
    for name in ('before', 'after'):
        directory = Path('output') / ('render-' + name)
        paths = list(directory.glob('*.png'))
        if len(paths) != 1:
            raise ValueError('Expected one sheet render')
        with Image.open(paths[0]) as im:
            if im.width != 780 or not 500 <= im.height <= 900:
                raise ValueError('Unexpected proof size: inspect sheet selection and print area')
            clean = Image.new('RGB', im.size, 'white')
            clean.paste(im.convert('RGB'))
        clean.save(Path('images') / (name + '.png'))
        pictures.append(clean)
    size = (max(p.width for p in pictures), max(p.height for p in pictures))
    frames = []
    for picture in pictures:
        frame = Image.new('RGB', size, 'white')
        frame.paste(picture, (0, 0))
        frames.append(frame)
    frames[1].save('images/cleanup.gif', save_all=True, append_images=[frames[0]], duration=[3200, 2200], loop=0)
    Path('evidence/images.json').write_text(json.dumps({
        'renderer': 'LibreOffice sheet export through xlrender; PNG rasterization with Poppler',
        'source_workbook': 'evidence/proof.xlsx', 'sheets': ['Before', 'After'],
        'pixels': list(size), 'frame_order': ['After', 'Before'], 'frame_milliseconds': [3200, 2200],
        'postprocessing': 'White canvas padding only; PNG metadata removed; two-frame GIF.'
    }, indent=2) + '\n')
    for name in ('before', 'after'):
        shutil.rmtree(Path('output') / ('render-' + name))
        (Path('output') / ('proof-' + name + '.xlsx')).unlink()


if __name__ == '__main__':
    main()
