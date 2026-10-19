# Build the PT-BR translation for the Switch version of Luminescent Platinum (BD), on top of the mod's own files.
# Source bundles: the mod's romfs first, else the game's romfs dump. Output: Output_Switch/romfs/... (merge it over the
# mod's romfs, it replaces the mod's Dpr.bin and message files).
# usage: python build_switch.py [mod AssetAssistant dir] [game dump AssetAssistant dir]
import os, sys, shutil, tempfile
import import_translation as I

MOD = r'E:\Downloads\Luminescent Platinum 2.2F BD-1-2-2F-1749583996\Luminescent Platinum 2.2F BD\romfs\Data\StreamingAssets\AssetAssistant'
DUMP = r'D:\Games\Emuladores\Citron\user\dump\0100000011D90000\romfs\Data\StreamingAssets\AssetAssistant'
OUT = os.path.join(I.HERE, 'Output_Switch', 'romfs', 'Data', 'StreamingAssets', 'AssetAssistant')

def main():
    mod, dump = (sys.argv[1:3] + [MOD, DUMP][len(sys.argv[1:3]):])
    rels = list(I.MESSAGES) + ['Dpr.bin']
    for d, dirs, files in os.walk(I.TEXTURES):
        dirs[:] = [x for x in dirs if x != 'sprites']
        rel = os.path.relpath(d, I.TEXTURES).replace('\\', '/')
        rels += [rel, rel + '.en']
    with tempfile.TemporaryDirectory() as tmp:   # the source files, mod first: the importer reads from one folder
        for rel in rels:
            src = next((p for p in (os.path.join(mod, rel), os.path.join(dump, rel)) if os.path.isfile(p)), None)
            if src:
                os.makedirs(os.path.dirname(os.path.join(tmp, rel)), exist_ok=True); shutil.copy(src, os.path.join(tmp, rel))
            elif rel in I.MESSAGES or rel == 'Dpr.bin': sys.exit('missing ' + rel)
        I.IN = tmp
        written = I.build_all(min_entries=1000)
    shutil.rmtree(os.path.join(I.HERE, 'Output_Switch'), ignore_errors=True)
    for rel, data in written.items():
        os.makedirs(os.path.dirname(os.path.join(OUT, rel)), exist_ok=True)
        open(os.path.join(OUT, rel), 'wb').write(data)
    print('Output_Switch: %d files' % len(written))

if __name__ == '__main__':
    main()
