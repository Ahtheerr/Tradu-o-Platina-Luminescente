# Build the PT-BR translation of Luminescent Platinum (PC port) from this folder.
#   luminescent_text.xlsx  the translation sheet (download it from Google Sheets as .xlsx over this file)
#   Textures/              edited images: a sheet PNG, or a sprite in sprites/ (same size as exported, or the
#                          sprite's full canvas size to give it more room)
#   Input/                 the untranslated game files this tool starts from (never modified)
#   Output/                the translated files, laid out like the game folder: copy its contents into the game
# usage: python import_translation.py        (needs: pip install UnityPy openpyxl pillow numpy)
# Sheet rows with bad codes keep the English text and are listed in luminescent_import_errors.txt.
import os, re, sys, copy, json, shutil, struct
import UnityPy, openpyxl
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
AAREL = os.path.join('Dpr_Data', 'StreamingAssets', 'AssetAssistant')
IN, OUT = os.path.join(HERE, 'Input', AAREL), os.path.join(HERE, 'Output', AAREL)
SHEET, TEXTURES = os.path.join(HERE, 'luminescent_text.xlsx'), os.path.join(HERE, 'Textures')
DATA = json.load(open(os.path.join(HERE, 'Input', 'tool_data.json'), encoding='utf-8'))
MESSAGES = ('Message/english', 'Message/common_msbt')

# ---- message text syntax (same as the sheet's Read me) ------------------------------------------------------------
GRM = 19                         # English grammar tag group: {male|female}, plurals, articles
ARTICLE = (3, 4, 5, 6)           # wordless grammar tags the game fills with an English "the"/"a" -> dropped
EV = {1: '\n', 3: '[p]', 4: '[s]'}
MARKUP = {'color': 2, 'size': 3, 'pos': 4}
METRIC = {'english_ss_zkn_height': 'spanish_ss_zkn_height',      # Pokédex sizes: the Spanish texts are metric
          'english_ss_zkn_weight': 'spanish_ss_zkn_weight'}      # in the pt-BR format ("0,7 m", "6,9 kg")
SHEETS = [   # sheet tab -> message file name parts (without english_); first match wins
    ('Dialogue', ['dp_scenario', 'dp_trainer_msg', 'dlp_rotom_message', 'dp_tvshow', 'dp_contest', 'dlp_underground',
                  'dlp_tsurearuki', 'ss_fld_shop', 'dp_poffin_main', 'ss_gameover', 'dlp_gmstation']),
    ('Battle', ['ss_btl_', 'ss_level_up', 'ss_shinka_demo', 'ss_tamago_demo', 'ss_waza_remember']),
    ('Descriptions', ['info', 'dp_pokedex_', 'dlp_adventure_note', 'ss_tips', 'ss_trainermemo']),
    ('Names', ['_name', 'ss_itemname', 'ss_wazaname', 'monsname', 'ss_tokusei', 'ss_seikaku', 'ss_typename',
               'dp_trainers_type', 'dp_characters', 'dp_fld_areaname', 'ss_ribbon', 'ss_boxname', 'ss_bag_pocket',
               'ss_zkn_type', 'ss_zkn_form']),
    ('Credits & Misc', ['staff_list', 'ss_zkn_height', 'ss_zkn_weight', 'halloffame', 'ss_illegalname']),
    ('Menus & UI', ['']),
]

def sheet_of(f):
    return next(name for name, parts in SHEETS if any(x in f[len('english_'):] for x in parts))

def fnum(v): return ('%g' % v) if v == round(v, 6) else repr(v)

def grm_param(tg, tags):
    P = tg['tagParameter']        # names a placeholder slot (tagIndex); shown as the [n] the text uses
    if P >= 253: return P         # 253 = party size, 254 = other trainer, 255 = player
    pos = [i for i, t in enumerate(tags) if t['tagIndex'] == P and t['groupID'] != GRM]
    return pos[0] if len(pos) == 1 else P

def to_text(lbl):
    """A label as sheet text (the 'english' column is this for the original label)."""
    tags, out = lbl['tagDataArray'], []
    for w in lbl['wordDataArray']:
        if w['patternID'] == 5:
            tg = tags[w['tagIndex']]
            if tg['groupID'] == GRM and tg['tagWordArray']:
                T, P = tg['tagID'], grm_param(tg, tags)
                pre = '' if (T, P) == (0, 255) else ('%d:' % P if T == 0 else '%d@%d:' % (T, P))
                out.append('{' + pre + '|'.join(tg['tagWordArray']) + '}')
            else: out.append('[%d]' % w['tagIndex'])
        else:
            out.append(w['str'].replace('[', '[[').replace('{', '{{'))
        ev = w['eventID']
        if ev in EV: out.append(EV[ev])
        elif ev == 2: out.append('[w:%s]' % fnum(w['tagValue']))
        elif ev == 5: out.append('[e:%s]' % fnum(w['tagValue']))
        elif ev not in (0, 7): raise ValueError('event %d' % ev)
    return ''.join(out)

TOK = re.compile(r'\[\[|\{\{|\n|\[p\]|\[s\]|\[w:([^\]]+)\]|\[e:([^\]]+)\]|\[(\d+)\]|\{(?:(\d+)(?:@(\d+))?:)?([^{}]*)\}|<(/?)(color|size|pos)[^>]*>')

def from_text(s):
    """-> list of (kind, value, eventID, tagValue); kind = 'txt' | 'tag' (tagDataArray index) | 'grm' ((T, P, words)) | 'mk'"""
    words, buf, pos = [], '', 0
    def close(ev, tv=0.0):
        nonlocal buf
        words.append(('txt', buf, ev, tv)); buf = ''
    for m in TOK.finditer(s):
        buf += s[pos:m.start()]; pos = m.end(); t = m.group(0)
        if t == '[[': buf += '['
        elif t == '{{': buf += '{'
        elif t == '\n': close(1)
        elif t == '[p]': close(3)
        elif t == '[s]': close(4)
        elif m.group(1): close(2, float(m.group(1)))
        elif m.group(2): close(5, float(m.group(2)))
        else:
            if buf: close(0)
            if m.group(3): words.append(('tag', int(m.group(3)), 0, 0.0))
            elif m.group(8): words.append(('mk', t, 0, 0.0))
            else:
                a, b = m.group(4), m.group(5)
                T, P = (int(a), int(b)) if b else (0, int(a) if a else 255)
                words.append(('grm', (T, P, tuple(m.group(6).split('|'))), 0, 0.0))
    buf += s[pos:]
    if not buf and (not words or words[-1][2] != 0): pass     # empty label, or one ending on an event
    elif buf or words[-1][2] != 0: close(7)
    else: words[-1] = words[-1][:2] + (7, 0.0)
    return words

ADV = DATA['glyph_advances']; AVG = sum(ADV.values()) / len(ADV)
def width(s):   # strWidth = sum of the font's horizontal advances
    return float(sum(ADV.get(c, AVG) for c in re.sub(r'<[^>]*>', '', s)))

def build(lbl, text):
    """New (wordDataArray, tagDataArray) for `text`; raises ValueError on bad codes."""
    tags = copy.deepcopy(lbl['tagDataArray'])
    is_grm = lambda t: t['groupID'] == GRM and t['tagWordArray']
    free = [i for i, t in enumerate(tags) if is_grm(t)]
    out = []
    for kind, v, ev, tv in from_text(text):
        w = {'patternID': 0, 'eventID': ev, 'tagIndex': -1, 'tagValue': tv, 'str': '', 'strWidth': -1.0}
        if kind == 'txt':
            if re.search(r'[\[\]{}]', v.replace('[[', '').replace('{{', '')):
                raise ValueError('unrecognised code near %r' % v)
            w['str'] = v
            w['patternID'] = 0 if ev == 0 else 7
            w['strWidth'] = width(v) if v else (-1.0 if ev == 7 else 0.0)
        elif kind == 'mk':
            w['str'] = v; w['patternID'] = MARKUP[re.match(r'</?(\w+)', v).group(1)]
        elif kind == 'tag':
            if v >= len(tags) or is_grm(tags[v]): raise ValueError('placeholder [%d] does not exist here' % v)
            if tags[v]['groupID'] == GRM and tags[v]['tagID'] in ARTICLE: continue
            w['patternID'] = 5; w['tagIndex'] = v
        else:
            T, P, ws = v
            if P < 253:
                if P >= len(tags) or is_grm(tags[P]): raise ValueError('{...} refers to missing placeholder [%d]' % P)
                P = tags[P]['tagIndex']
            entry = {'tagIndex': -1, 'groupID': GRM, 'tagID': T, 'tagPatternID': 5, 'forceArticle': 0,
                     'tagParameter': P, 'tagWordArray': list(ws), 'forceGrmID': 0}
            if free: pos = free.pop(0); tags[pos] = entry
            else: pos = len(tags); tags.append(entry)
            w['patternID'] = 5; w['tagIndex'] = pos
        out.append(w)
    # End the message the way the original does: its last word is a text word whose event (7 = end, sometimes
    # 0/1/...) and patternID (0 or 7, per message, no rule predicts it) we copy. A wrong ending leaves some windows
    # (signs) unable to close. A placeholder/markup never carries the end: an empty text word follows it instead.
    old = lbl['wordDataArray'][-1] if lbl['wordDataArray'] else None
    otext = old is not None and old['patternID'] in (0, 7)
    oev, opat = (old['eventID'], old['patternID']) if otext else (7, 7)
    last = out[-1] if out else None
    if last and last['patternID'] in (0, 7) and last['eventID'] in (0, 7, oev):
        last['eventID'], last['patternID'] = oev, opat
    elif not otext and old and last and last['patternID'] == old['patternID'] and last['eventID'] == old['eventID'] == 7:
        pass                                                      # original also ends on this placeholder/markup
    else:
        out.append({'patternID': opat if oev == 7 else 7, 'eventID': 7, 'tagIndex': -1, 'tagValue': 0.0, 'str': '', 'strWidth': -1.0})
    return out, tags

def translations(path):
    """(sheet tab, english) -> translation. Formula cells give their saved result, which only exists if Google
    Sheets/Excel saved the file."""
    out = {}
    raw = openpyxl.load_workbook(path, read_only=True)
    for sh in openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets:
        rows = sh.iter_rows(values_only=True)
        hdr = next(rows, None)
        if not hdr or 'translation' not in hdr or 'english' not in hdr: continue
        e, t = hdr.index('english'), hdr.index('translation')
        frows = raw[sh.title].iter_rows(values_only=True); next(frows)
        for r, f in zip(rows, frows):
            if r[t] is None and f[t] is not None:
                sys.exit('%s: formulas have no saved results (sheet %s). Open the file in Google Sheets/Excel and download/save it again.' % (path, sh.title))
            if r[t] not in (None, '') and r[e]: out[(sh.title, r[e])] = str(r[t]).replace('\r\n', '\n')
    return out

def import_messages(written):
    tr = translations(SHEET)
    done, errors = 0, []
    for rel in MESSAGES:
        env = UnityPy.load(os.path.join(IN, rel))
        metric = {}
        for o in env.objects:
            if o.type.name == 'MonoBehaviour' and o.peek_name() in METRIC.values():
                metric[o.peek_name()] = {l['labelName']: to_text(l) for l in o.read_typetree()['labelDataArray']}
        for o in env.objects:
            if o.type.name != 'MonoBehaviour' or not o.peek_name().startswith('english_'): continue
            t = o.read_typetree()
            sheet = sheet_of(t['m_Name']); changed = False
            for lbl in t['labelDataArray']:
                text = tr.get((sheet, to_text(lbl)))
                if t['m_Name'] in METRIC: text = metric[METRIC[t['m_Name']]].get(lbl['labelName']) or text
                if text is None: continue
                try:
                    lbl['wordDataArray'], lbl['tagDataArray'] = build(lbl, text)
                    done += 1; changed = True
                except ValueError as ex:
                    errors.append('%s\t%s\t%s\t%r\t%s' % (sheet, t['m_Name'], lbl['labelName'], text, ex))
            if changed: o.save_typetree(t)
        written[rel] = list(env.files.values())[0].save(packer='lz4')
    with open(os.path.join(HERE, 'luminescent_import_errors.txt'), 'w', encoding='utf-8') as fh:
        fh.write('sheet\tfile\tlabel\ttranslation\tproblem\n' + '\n'.join(errors) + '\n')
    print('text: %d game entries from %d sheet rows; %d kept English because of errors (see luminescent_import_errors.txt)'
          % (done, len(tr), len(errors)))
    return done

# ---- textures --------------------------------------------------------------------------------------------------------
def free_spot(sheet, size, used):
    """Top-left (x, y) of a fully transparent area of `size` (+2 px margin, on a 6 px grid) in `sheet` that no sprite
    rect in `used` covers, or None."""
    import numpy as np
    busy = np.array(sheet)[:, :, 3] > 0
    for r in used:                                                # transparent corners of other sprites are theirs too
        y0 = sheet.height - int(r['y'] + r['height']) - 1
        busy[max(y0, 0):sheet.height - int(r['y']) + 1, max(int(r['x']) - 1, 0):int(r['x'] + r['width']) + 2] = True
    w, h = size[0] + 4, size[1] + 4
    ii = np.pad(busy.cumsum(0).cumsum(1), ((1, 0), (1, 0)))
    for y in range(0, sheet.height - h + 1, 6):
        for x in range(0, sheet.width - w + 1, 6):
            if not ii[y + h, x + w] - ii[y, x + w] - ii[y + h, x] + ii[y, x]: return x + 2, y + 2
    return None

def import_textures(d, rel):
    """Apply Textures/<d> to the Input bundle `rel`; -> new bundle bytes, or None if nothing changed."""
    pngs = {os.path.splitext(f)[0]: os.path.join(d, f) for f in os.listdir(d) if f.endswith('.png')}
    env = UnityPy.load(os.path.join(IN, rel))
    sheets = {}                                                   # texture path id -> (Texture2D, edited, original)
    for o in env.objects:
        if o.type.name != 'Texture2D': continue
        t = o.read(); img = t.image.convert('RGBA')
        p = pngs.get(t.m_Name.replace(':', '_').replace('/', '_'))
        new = Image.open(p).convert('RGBA') if p else img
        if new.size != img.size:
            print('SKIP %s: %s, must stay %dx%d' % (p, new.size, *img.size)); new = img
        sheets[o.path_id] = (t, new, img)
    # sprites/: a crop of the original size goes back in place; a PNG of the sprite's full canvas (m_Rect) is trimmed
    # and moved to free space in its sheet, so the text can grow
    sp_dir = os.path.join(d, 'sprites')
    key = lambda k: (tuple(k[0].values()), k[1])
    atlases = [(o, o.read_typetree()) for o in env.objects if o.type.name == 'SpriteAtlas']
    atlas = {key(k): v for _, a in atlases for k, v in a['m_RenderDataMap']}
    sprites = [(o, o.read_typetree()) for o in env.objects if o.type.name == 'Sprite']
    rds = lambda s: [r for r in (s['m_RD'], atlas.get(key(s['m_RenderDataKey']))) if r and r['texture']['m_PathID']]
    moved = False
    for o, s in sprites if os.path.isdir(sp_dir) else ():
        p = os.path.join(sp_dir, s['m_Name'].replace(':', '_').replace('/', '_') + '.png')
        if not os.path.exists(p): continue
        rd = (rds(s) or [None])[-1]                               # the SpriteAtlas entry when there is one
        if not rd or rd['texture']['m_PathID'] not in sheets: print('SKIP %s: sheet not found' % p); continue
        t, new, img = sheets[rd['texture']['m_PathID']]
        r = rd['textureRect']; size = (round(r['width']), round(r['height']))
        sp = Image.open(p).convert('RGBA')
        if sp.tobytes() == o.read().image.convert('RGBA').tobytes(): continue
        # ponytail: only plain rectangle-packed, unrotated sprites (settingsRaw bits 1..5); tight/rotated ones are skipped
        full = (round(s['m_Rect']['width']), round(s['m_Rect']['height']))
        if rd['settingsRaw'] & 0x3e != 2 or sp.size not in (size, full):
            print('SKIP %s: rotated/tight-packed or wrong size (%dx%d crop or %dx%d full canvas)' % (p, *size, *full)); continue
        if sp.size == size:
            new.paste(sp, (round(r['x']), img.height - round(r['y'] + r['height']))); continue
        l, tp, rr, b = sp.getbbox() or (0, 0, 1, 1)
        l, tp, rr, b = max(l - 1, 0), max(tp - 1, 0), min(rr + 1, sp.width), min(b + 1, sp.height)   # 1 px clear border
        crop = sp.crop((l, tp, rr, b))
        others = [x['textureRect'] for o2, s2 in sprites if o2 is not o for x in rds(s2)
                  if x['texture']['m_PathID'] == rd['texture']['m_PathID']]
        pos = free_spot(new, crop.size, others)
        if not pos: print('SKIP %s: no free space in the sheet' % p); continue
        X, Y = pos; new.paste(crop, (X, Y))
        ty = img.height - Y - crop.height                         # Unity rects count y from the bottom
        ox, oy = l, sp.height - b
        for x in rds(s):
            if x['texture']['m_PathID'] != rd['texture']['m_PathID']: continue
            x['textureRect'] = {'x': float(X), 'y': float(ty), 'width': float(crop.width), 'height': float(crop.height)}
            x['textureRectOffset'] = {'x': float(ox), 'y': float(oy)}
            x['atlasRectOffset'] = {'x': float(X - ox), 'y': float(ty - oy)}
            x['uvTransform']['y'] = X - ox + full[0] * s['m_Pivot']['x']
            x['uvTransform']['w'] = ty - oy + full[1] * s['m_Pivot']['y']
        # ponytail: the sprite mesh (m_VertexData) keeps the old outline; UI Images don't use it, SpriteRenderers would
        o.save_typetree(s); moved = True
    if moved:
        for o, a in atlases: o.save_typetree(a)
    changed = []
    for t, new, img in sheets.values():
        if new.tobytes() == img.tobytes(): continue
        # ponytail: Switch (platform 38) gets one mip level, UnityPy swizzles only the first; fine for UI, 3D objects
        # lose their smaller mips
        mips = 1 if t.object_reader.platform == 38 else (t.m_MipCount or 1)
        t.set_image(new, target_format=t.m_TextureFormat, mipmap_count=mips)
        t.save(); changed.append(t.m_Name)
    if not changed: return None
    print('textures: %s %s' % (rel, changed))
    return list(env.files.values())[0].save(packer='lz4')

def dpr_size_offsets(data):
    """Dpr.bin (AssetBundleDownloadManifest) -> {bundle path: file offset of its size field}; read from the game's own
    file so any build works."""
    import nrbf
    rd = nrbf.Reader(data); rd.read_all()
    return {nrbf.deref(rd, o['projectName']) + '/' + nrbf.deref(rd, o['assetBundleName']): o['size'][2]
            for o in rd.objects.values() if isinstance(o, dict) and o['_class'].endswith('AssetBundleRecord')}

def texture_bundles():
    """[(Textures/ subfolder, bundle path in IN)] for every folder with PNGs."""
    out = []
    for d, dirs, files in os.walk(TEXTURES):
        dirs[:] = [x for x in dirs if x != 'sprites']
        if not any(f.endswith('.png') for f in files): continue
        rel = os.path.relpath(d, TEXTURES).replace('\\', '/')
        rel = next((r for r in (rel + '.en', rel) if os.path.isfile(os.path.join(IN, r))), None)
        if not rel: print('SKIP %s: its bundle is not in the game files' % d); continue
        out.append((d, rel))
    return out

def build_all(min_entries=0):
    """-> {path in AssetAssistant: new file bytes}, Dpr.bin included. Stops if fewer than `min_entries` texts match
    (the game's text is not the English this sheet was made from, e.g. already translated)."""
    written = {}
    if import_messages(written) < min_entries:
        raise RuntimeError('The game text does not match the sheet (already translated, or not Luminescent Platinum?)')
    for d, rel in texture_bundles():
        data = import_textures(d, rel)
        if data: written[rel] = data
    # Dpr.bin lists every bundle's size: patch the ones we wrote
    dpr = bytearray(open(os.path.join(IN, 'Dpr.bin'), 'rb').read())
    offs = dpr_size_offsets(dpr)
    for rel, data in written.items():
        struct.pack_into('<q', dpr, offs[rel], len(data))
    written['Dpr.bin'] = bytes(dpr)
    return written

def main():
    written = build_all()
    shutil.rmtree(os.path.join(HERE, 'Output'), ignore_errors=True)
    for rel, data in written.items():
        os.makedirs(os.path.dirname(os.path.join(OUT, rel)), exist_ok=True)
        open(os.path.join(OUT, rel), 'wb').write(data)
    print('Output: %d files. Copy the contents of Output into the game folder.' % len(written))

if __name__ == '__main__':
    main()
