# -*- coding: utf-8 -*-

import os
import csv
import re
import sys
from bisect import bisect_right
import music21
import copy

# Optional Excel support
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

# ---------------------------
# Logger Class for Debug Files
# ---------------------------

class Logger:
    """A simple logger class to write to both console and a file."""
    def __init__(self, log_path):
        self.terminal = sys.stdout
        self.log_file = open(log_path, 'w', encoding='utf-8')

    def write(self, message):
        self.terminal.write(message)
        self.log_file.write(message)

    def flush(self):
        self.terminal.flush()
        self.log_file.flush()

    def close(self):
        self.log_file.close()

# ---------------------------
# Utilities
# ---------------------------

def preprocess_musicxml(xml_content: str) -> str:
    """
    Finds and corrects non-standard harmony tags by modifying the root tag
    based on the kind text before parsing. This is the robust solution.
    """
    # This pattern is more advanced. It finds a <root-step> and then looks ahead
    # for a related <kind text="...">other</kind> tag. It captures three groups:
    # 1. The full <root-step> tag itself.
    # 2. Everything between the root-step and the kind tag.
    # 3. The text content of the kind tag.
    pattern = re.compile(r'(<root-step>[A-G]</root-step>)(.*?)<kind\s+text="([^"]+)">other</kind>', re.DOTALL)

    def replacer(match):
        root_step_tag = match.group(1)  # e.g., "<root-step>E</root-step>"
        intervening_xml = match.group(2) # e.g., newlines and spacing
        kind_text = match.group(3).strip() # e.g., "♭sus"

        new_root_alter = ""
        new_kind_tag = '<kind text="">major</kind>'  # Default fallback

        # Check for accidentals in the kind text and create a root-alter tag
        if '♭' in kind_text or 'b' in kind_text:
            new_root_alter = "<root-alter>-1</root-alter>"
        elif '#' in kind_text:
            new_root_alter = "<root-alter>1</root-alter>"

        # Determine the correct new kind tag
        if 'sus' in kind_text:
            new_kind_tag = '<kind text="sus4">suspended-fourth</kind>'
        elif 'add9' in kind_text:
            new_kind_tag = '<kind>add-9</kind>'

        # Reconstruct the XML snippet correctly
        # This inserts the new <root-alter> tag right after the <root-step> tag
        return f"{root_step_tag}{new_root_alter}{intervening_xml}{new_kind_tag}"

    return pattern.sub(replacer, xml_content)


def split_tone_cell(cell: str) -> list[str]:
    """Split a cell like 'C - E - G' into a clean list ['C', 'E', 'G']."""
    return [tone.strip() for tone in str(cell).split('-') if tone.strip()]

def safe_cell(x):
    """Coerce any cell into a clean string."""
    if x is None:
        return ''
    return str(x)

# ---------------------------
# Table loader
# ---------------------------

def find_header_index(rows_2d):
    for i, row in enumerate(rows_2d):
        if not row:
            continue
        first = safe_cell(row[0]).strip()
        if first.startswith('Chord Symbol'):
            later = {safe_cell(c).strip() for c in row[1:] if safe_cell(c).strip()}
            if any(t in later for t in ('X', 'Xm', 'Xmaj7', 'Xm7')):
                return i
    return -1

def load_chord_table(path, sheet_name='detailed_chord_table_pretty'):
    rows = None
    ext = os.path.splitext(path)[1].lower()
    if HAS_PANDAS and ext in ('.xlsx', '.xlsm', '.xlsb', '.ods'):
        try:
            df = pd.read_excel(path, header=None, sheet_name=sheet_name, engine='openpyxl')
            rows = df.fillna('').astype(str).values.tolist()
        except Exception as e:
            print(f"Warning: Excel load failed ({e}); trying CSV fallback.")

    if rows is None:
        try:
            with open(path, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                rows = [list(map(safe_cell, r)) for r in reader]
        except FileNotFoundError:
            print(f"Error: The file at {path} was not found.")
            return None
        except Exception as e:
            print(f"An error occurred while loading the chord table: {e}")
            return None

    hdr_idx = find_header_index(rows)
    if hdr_idx < 0:
        print("Error: Could not locate the 'Chord Symbol - X = Note name' header row.")
        return None

    header = rows[hdr_idx]
    templates = {i: safe_cell(tmpl).strip() for i, tmpl in enumerate(header) if i > 0 and safe_cell(tmpl).strip()}

    chord_dict = {}
    for r in rows[hdr_idx + 1:]:
        if not r or all(safe_cell(c).strip() == '' for c in r):
            continue
        root_raw = safe_cell(r[0]).strip()
        if not root_raw:
            continue

        for col_idx, tmpl in templates.items():
            if col_idx >= len(r):
                continue

            key = tmpl.replace('X', root_raw)
            tones = split_tone_cell(safe_cell(r[col_idx]))

            if key and tones and key not in chord_dict:
                chord_dict[key] = tones
    return chord_dict

def get_lookup_key_from_figure(figure_str: str) -> str:
    """
    Directly converts a chord's text name (figure) into the correct
    lookup key by handling common variations in chord notation.
    This version includes the fix for the standalone "C-" issue.
    """
    # Enhanced slash chord logic
    if '/' in figure_str:
        parts = figure_str.split('/', 1)
        if parts[1].strip().startswith(('M', 'm', 's', 'd', 'a', 'p')):
            key = ''.join(parts)
        else:
            key = parts[0]
    else:
        key = figure_str

    key = key.replace('♭', 'b')

    # --- SOLUTION: New fix for the standalone "C-" major chord case ---
    # This specifically handles cases where the figure is exactly "C-",
    # which music21 may produce for a C major chord with a "-" kind text.
    if key == 'C-':
        return 'C'

    # Targeted fix for the unique "C-" case (now correctly handles C-7 etc.)
    if key.startswith('C-'):
        if len(key) == 2:
            # This case is now handled by the check above, but we leave
            # this structure in case of other edge cases.
            key = 'B'
        else:
            key = 'B' + key[2:]

    # UNIFIED LOGIC for all other flat chords
    elif key.startswith(('A-', 'B-', 'D-', 'E-', 'G-')):
        root = key[0] + 'b'
        quality = key[2:]
        key = root + quality

    # Original logic for F-flat
    elif key.startswith('F-'):
        key = 'E' + key[2:]

    # General cleanup and standardization
    key = key.replace(' ', '').replace('(', '').replace(')', '')
    key = key.replace('power', '5')
    if 'mM' in key:
        key = key.replace('mM', 'mMaj')
    key = key.replace('-maj', 'maj').replace('-M', 'maj')
    key = key.replace('-', '')
    if key.endswith('sus'):
        if not key.endswith('sus2') and not key.endswith('sus4'):
            key += '4'
    if 'm' not in key and 'M' in key:
        key = key.replace('M', 'maj')

    return key

def find_closest_harmony_note(melody_note, chord_symbol, chord_table):
    key = get_lookup_key_from_figure(chord_symbol.figure)

    measure_num = melody_note.measureNumber
    lyric = melody_note.lyric if melody_note.lyric else "No Lyric"

    print(f"\n--- M.{measure_num} | Lyric: '{lyric}' ---")
    print(f"   -> Melody: '{melody_note.nameWithOctave}', Chord: '{chord_symbol.figure}', Mapped to key: '{key}'")

    if key not in chord_table:
        print(f"       - ⚠️ Warning: Key '{key}' not found in chord table. No harmony will be generated.")
        return None

    tones = chord_table[key]
    print(f"       - Available Tones in Table: {tones}")

    mel_oct = melody_note.pitch.octave
    candidates = []
    for name in tones:
        for octv in range(max(0, mel_oct - 1), max(0, mel_oct) + 3):
            try:
                candidates.append(music21.pitch.Pitch(f"{name}{octv}"))
            except music21.pitch.PitchException:
                pass

    if not candidates:
        return None

    upward_candidates = []
    for p in candidates:
        intv = music21.interval.Interval(melody_note.pitch, p)
        if intv.direction > 0 and intv.semitones >= 3:
            upward_candidates.append((p, intv.semitones))

    if upward_candidates:
        sorted_candidates = sorted(upward_candidates, key=lambda t: t[1])
        valid_notes_str = ', '.join([p.nameWithOctave for p, s in sorted_candidates])
        print(f"       - Valid Options (>= 3 semitones above): [{valid_notes_str}]")

        pitch_to_return, _ = sorted_candidates[0]
        print(f"       - ✅ Selection: Chose '{pitch_to_return.nameWithOctave}' (it's the closest valid option).")
        return pitch_to_return

    print(f"       - ℹ️ Info: No harmony note found (all available tones were less than 3 semitones away).")
    return None

# ---------------------------
# Core processing
# ---------------------------
def generate_harmony_for_file(melody_path, chord_table):
    try:
        with open(melody_path, 'r', encoding='utf-8') as f:
            xml_text = f.read()

        # --- NEW TARGETED FIX for 'Lately' XML issue ---
        # This corrects a specific non-standard tag for minor-major-seventh chords
        # before music21 parsing, ensuring the quality is not lost.
        xml_text = xml_text.replace('<kind text="m(maj7)">other</kind>', '<kind>minor-major-seventh</kind>')

        corrected_xml_text = preprocess_musicxml(xml_text)
        score = music21.converter.parse(corrected_xml_text)
    except Exception as e:
        print(f"Error parsing melody file: {melody_path}. Error: {e}")
        return
    best_part = None
    if score.parts:
        best_part = score.parts[0]
    for part in score.parts:
        if any(n.lyric for n in part.recurse().notes):
            best_part = part
            break

    if not best_part:
        print("       - Error: No parts found in the score. Skipping.")
        return

    print(f"       - Using part '{best_part.partName or 'Unnamed Part'}' as melody.")

    harmony_part = music21.stream.Part(id='harmony', partName='Harmony')

    for item in best_part.flatten().getElementsByClass(['Clef', 'KeySignature', 'TimeSignature']):
        if item.offset == 0.0:
          harmony_part.insert(item.offset, copy.deepcopy(item))

    chords = sorted(score.flatten().getElementsByClass('ChordSymbol'), key=lambda cs: cs.offset)
    ch_offs = [float(cs.offset) for cs in chords]

    for element in best_part.flatten().notesAndRests:
        new_element = None
        if isinstance(element, music21.note.Note):
            active_chord = chords[bisect_right(ch_offs, float(element.offset)) - 1] if ch_offs else None
            if active_chord:
                harmony_pitch = find_closest_harmony_note(element, active_chord, chord_table)
                if harmony_pitch:
                    new_element = music21.note.Note(harmony_pitch, duration=element.duration)
                    if element.lyric:
                        new_element.lyric = element.lyric
                else:
                    new_element = music21.note.Rest(duration=element.duration)
            else:
                new_element = music21.note.Rest(duration=element.duration)
        elif isinstance(element, music21.note.Rest):
            new_element = music21.note.Rest(duration=element.duration)

        if new_element:
            harmony_part.insert(element.offset, new_element)

    input_dir = os.path.dirname(melody_path)
    song_dir = os.path.dirname(input_dir)
    output_dir = os.path.join(song_dir, 'output')
    base_name = os.path.splitext(os.path.basename(melody_path))[0]
    output_xml_path = os.path.join(output_dir, f'harmony_output_{base_name}.xml')

    out_score = music21.stream.Score()
    out_score.insert(0, harmony_part)
    out_score.makeNotation(inPlace=True)

    out_score.write('musicxml', fp=output_xml_path)
    print(f"\n       - Successfully generated harmony and saved to '{output_xml_path}'")

# ---------------------------
# Batch Processor
# ---------------------------
def process_all_songs(root_directory, chord_table_path):
    print("--- Starting Batch Harmony Generation ---\n")
    chord_table = load_chord_table(chord_table_path)
    if chord_table is None:
        print("Could not load chord table. Halting process.")
        return

    for song_name in os.listdir(root_directory):
        song_dir = os.path.join(root_directory, song_name)
        if not os.path.isdir(song_dir):
            continue

        input_dir = os.path.join(song_dir, 'input')
        output_dir = os.path.join(song_dir, 'output')

        if not os.path.isdir(input_dir):
            continue

        print(f"Processing song folder: '{song_name}'")
        os.makedirs(output_dir, exist_ok=True)

        for filename in os.listdir(input_dir):
            if filename.lower().endswith(('.xml', '.musicxml')):
                if 'harmony_output' in filename:
                    continue

                melody_file_path = os.path.join(input_dir, filename)
                base_name = os.path.splitext(filename)[0]
                log_path = os.path.join(output_dir, f'harmony_output_{base_name}_log.txt')

                original_stdout = sys.stdout
                sys.stdout = Logger(log_path)

                try:
                    print(f"--- Generating Harmony for: {filename} ---")
                    generate_harmony_for_file(melody_file_path, chord_table)
                except Exception as e:
                    print("\n--- ERROR ---")
                    print(f"An unexpected error occurred while processing {filename}.")
                    print(f"Error details: {e}")
                finally:
                    if isinstance(sys.stdout, Logger):
                        sys.stdout.close()
                    sys.stdout = original_stdout
                    print(f"     - Finished processing '{filename}'. Log saved to '{log_path}'")

    print("\n--- Batch Processing Complete ---")

# --- Run ---
if __name__ == '__main__':
    SONGS_DIRECTORY = 'Songs_v8'
    CHORD_TABLE_FILE = '/content/detailed_chord_table.xlsx'

    if not os.path.exists(CHORD_TABLE_FILE):
        print(f"Info: '{CHORD_TABLE_FILE}' not found. Looking for CSV version.")
        CHORD_TABLE_FILE = 'detailed_chord_table_pretty.csv'

    process_all_songs(SONGS_DIRECTORY, CHORD_TABLE_FILE)