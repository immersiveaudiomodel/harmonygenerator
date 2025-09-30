# -*- coding: utf-8 -*-
"""
Created on Sat Sep 13 12:50:22 2025
"""
import os
import csv
import re
import sys
import copy
import gradio as gr
import music21
import subprocess
import tempfile
from contextlib import redirect_stdout
import io
from bisect import bisect_right

# --- Configuration ---
# Set the path to your SoundFont file.
# The script expects it to be in the same directory.
SOUNDFONT_PATH = "GeneralUser_GS_v1.471.sf2"
CHORD_TABLE_FILE = 'detailed_chord_table.xlsx' # Or 'detailed_chord_table_pretty.csv'

# Optional Excel support
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


# ---------------------------
# Logger Class for Debug Files (No longer used, but kept for reference)
# ---------------------------
class Logger:
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
    pattern = re.compile(r'(<root-step>[A-G]</root-step>)(.*?)<kind\s+text="([^"]+)">other</kind>', re.DOTALL)
    def replacer(match):
        root_step_tag, intervening_xml, kind_text = match.groups()
        kind_text = kind_text.strip()
        new_root_alter = ""
        new_kind_tag = '<kind text="">major</kind>'
        if '?' in kind_text or 'b' in kind_text: new_root_alter = "<root-alter>-1</root-alter>"
        elif '#' in kind_text: new_root_alter = "<root-alter>1</root-alter>"
        if 'sus' in kind_text: new_kind_tag = '<kind text="sus4">suspended-fourth</kind>'
        elif 'add9' in kind_text: new_kind_tag = '<kind>add-9</kind>'
        return f"{root_step_tag}{new_root_alter}{intervening_xml}{new_kind_tag}"
    return pattern.sub(replacer, xml_content)

def split_tone_cell(cell: str) -> list[str]:
    return [tone.strip() for tone in str(cell).split('-') if tone.strip()]

def safe_cell(x):
    return '' if x is None else str(x)

# ---------------------------
# Table loader
# ---------------------------
def find_header_index(rows_2d):
    for i, row in enumerate(rows_2d):
        if not row: continue
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
        if not r or all(safe_cell(c).strip() == '' for c in r): continue
        root_raw = safe_cell(r[0]).strip()
        if not root_raw: continue
        for col_idx, tmpl in templates.items():
            if col_idx >= len(r): continue
            key = tmpl.replace('X', root_raw)
            tones = split_tone_cell(safe_cell(r[col_idx]))
            if key and tones and key not in chord_dict:
                chord_dict[key] = tones
    return chord_dict

# ---------------------------
# Chord Mapping & Harmony Selection
# ---------------------------
def get_lookup_key_from_figure(figure_str: str) -> str:
    if '/' in figure_str:
        parts = figure_str.split('/', 1)
        key = ''.join(parts) if parts[1].strip().startswith(('M', 'm', 's', 'd', 'a', 'p')) else parts[0]
    else:
        key = figure_str
    key = key.replace('?', 'b')
    if key == 'C-': return 'C'
    if key.startswith('C-'): key = 'B' if len(key) == 2 else 'B' + key[2:]
    elif key.startswith(('A-', 'B-', 'D-', 'E-', 'G-')): key = key[0] + 'b' + key[2:]
    elif key.startswith('F-'): key = 'E' + key[2:]
    key = key.replace(' ', '').replace('(', '').replace(')', '')
    key = key.replace('power', '5')
    if 'mM' in key: key = key.replace('mM', 'mMaj')
    key = key.replace('-maj', 'maj').replace('-M', 'maj')
    key = key.replace('-', '')
    if key.endswith('sus') and not key.endswith(('sus2', 'sus4')): key += '4'
    if 'm' not in key and 'M' in key: key = key.replace('M', 'maj')
    return key

def find_closest_harmony_note(melody_note, chord_symbol, chord_table):
    key = get_lookup_key_from_figure(chord_symbol.figure)
    measure_num = melody_note.measureNumber
    lyric = melody_note.lyric if melody_note.lyric else "No Lyric"
    print(f"\n--- M.{measure_num} | Lyric: '{lyric}' ---")
    print(f"   -> Melody: '{melody_note.nameWithOctave}', Chord: '{chord_symbol.figure}', Mapped to key: '{key}'")
    if key not in chord_table:
        print(f"      - ⚠️ Warning: Key '{key}' not found in chord table. No harmony will be generated.")
        return None
    tones = chord_table[key]
    print(f"      - Available Tones in Table: {tones}")
    mel_oct = melody_note.pitch.octave
    candidates = []
    for name in tones:
        for octv in range(max(0, mel_oct - 1), max(0, mel_oct) + 3):
            try:
                candidates.append(music21.pitch.Pitch(f"{name}{octv}"))
            except music21.pitch.PitchException: pass
    if not candidates: return None
    upward_candidates = []
    for p in candidates:
        intv = music21.interval.Interval(melody_note.pitch, p)
        if intv.direction > 0 and intv.semitones >= 3:
            upward_candidates.append((p, intv.semitones))
    if upward_candidates:
        sorted_candidates = sorted(upward_candidates, key=lambda t: t[1])
        valid_notes_str = ', '.join([p.nameWithOctave for p, s in sorted_candidates])
        print(f"      - Valid Options (>= 3 semitones above): [{valid_notes_str}]")
        pitch_to_return, _ = sorted_candidates[0]
        print(f"      - ✅ Selection: Chose '{pitch_to_return.nameWithOctave}' (it's the closest valid option).")
        return pitch_to_return
    print(f"      - ℹ️ Info: No harmony note found (all available tones were less than 3 semitones away).")
    return None

# ---------------------------
# Core processing
# ---------------------------
def generate_harmony_score(melody_path, chord_table):
    """
    Modified version of your function.
    It now returns a music21.stream.Score object instead of writing a file.
    """
    try:
        with open(melody_path, 'r', encoding='utf-8') as f:
            xml_text = f.read()
        xml_text = xml_text.replace('<kind text="m(maj7)">other</kind>', '<kind>minor-major-seventh</kind>')
        corrected_xml_text = preprocess_musicxml(xml_text)
        score = music21.converter.parse(corrected_xml_text)
    except Exception as e:
        print(f"Error parsing melody file: {melody_path}. Error: {e}")
        return None

    best_part = score.parts[0] if score.parts else None
    for part in score.parts:
        if any(n.lyric for n in part.recurse().notes):
            best_part = part
            break
    if not best_part:
        print("Error: No parts with lyrics found in the score. Using first part.")
        if not score.parts:
             print("Error: No parts found in score. Cannot proceed.")
             return None
        best_part = score.parts[0]

    print(f"Using part '{best_part.partName or 'Unnamed Part'}' as melody.")
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
                    if element.lyric: new_element.lyric = element.lyric
                else: new_element = music21.note.Rest(duration=element.duration)
            else: new_element = music21.note.Rest(duration=element.duration)
        elif isinstance(element, music21.note.Rest):
            new_element = music21.note.Rest(duration=element.duration)
        if new_element:
            harmony_part.insert(element.offset, new_element)

    out_score = music21.stream.Score()
    out_score.insert(0, harmony_part)
    out_score.makeNotation(inPlace=True)
    return out_score

# ==============================================================================
#  NEW GRADIO-SPECIFIC AND AUDIO CONVERSION CODE
# ==============================================================================

def synthesize_midi_to_wav(midi_path, wav_path):
    """Converts a MIDI file to a WAV file using FluidSynth."""
    if not os.path.exists(SOUNDFONT_PATH):
        print("ERROR: SoundFont file not found at:", SOUNDFONT_PATH)
        print("Please download a .sf2 file and place it in the correct path.")
        return False
    try:
        command = [
            'fluidsynth', '-ni', SOUNDFONT_PATH, midi_path,
            '-F', wav_path, '-r', '44100'
        ]
        #print(f"\nRunning FluidSynth command: {' '.join(command)}")
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        #print("Successfully synthesized WAV file.")
        return True
    except FileNotFoundError:
        print("\n---")
        print("ERROR: `fluidsynth` command not found.")
        print("Please install FluidSynth on your system and ensure it's in the system's PATH.")
        print("---")
        return False
    except subprocess.CalledProcessError as e:
        print(f"\n---")
        print(f"ERROR: FluidSynth failed to convert MIDI to WAV.")
        print(f"Stderr: {e.stderr.decode()}")
        print(f"---")
        return False
    except Exception as e:
        print(f"An unexpected error occurred during WAV synthesis: {e}")
        return False


def generate_harmony_and_outputs(input_xml_file):
    """
    This is the main function called by the Gradio interface.
    It orchestrates the harmony generation and file conversions.
    """
    if input_xml_file is None:
        return None, None, None, "Please upload a MusicXML file or select a sample."

    # Capture all print outputs as logs
    log_stream = io.StringIO()
    with redirect_stdout(log_stream):
        try:
            print("--- Starting Harmony Generation Process ---")
            
            # 1. Setup paths
            input_path = input_xml_file.name
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            temp_dir = tempfile.mkdtemp()
            output_xml_path = os.path.join(temp_dir, f"harmony_output_{base_name}.xml")
            output_mid_path = os.path.join(temp_dir, f"harmony_output_{base_name}.mid")
            output_wav_path = os.path.join(temp_dir, f"harmony_output_{base_name}.wav")

            print(f"Processing input file: {os.path.basename(input_path)}")
            
            # 2. Load Chord Table
            print(f"Loading chord table from: {CHORD_TABLE_FILE}")
            chord_table = load_chord_table(CHORD_TABLE_FILE)
            if chord_table is None:
                raise Exception("Failed to load chord table. Halting process.")

            # 3. Generate Harmony Score
            harmony_score = generate_harmony_score(input_path, chord_table)
            if harmony_score is None:
                raise Exception("Failed to generate a harmony score object.")

            # 4. Write output files (XML and MIDI)
            #print(f"\n--- Writing Output Files ---")
            harmony_score.write('musicxml', fp=output_xml_path)
            #print(f"Successfully wrote MusicXML to: {os.path.basename(output_xml_path)}")
            harmony_score.write('midi', fp=output_mid_path)
            #print(f"Successfully wrote MIDI to: {os.path.basename(output_mid_path)}")
            
            # 5. Synthesize WAV from MIDI
            wav_success = synthesize_midi_to_wav(output_mid_path, output_wav_path)
            
            print("\n--- Process Complete ---")

            # Prepare outputs for Gradio
            final_wav_path = output_wav_path if wav_success and os.path.exists(output_wav_path) else None
            logs = log_stream.getvalue()
            return output_xml_path, output_mid_path, final_wav_path, logs

        except Exception as e:
            print("\n--- A CRITICAL ERROR OCCURRED ---")
            print(str(e))
            logs = log_stream.getvalue()
            return None, None, None, logs

# --- Create samples folder if it doesn't exist ---
SAMPLES_DIR = "samples"
os.makedirs(SAMPLES_DIR, exist_ok=True)
sample_files = [os.path.join(SAMPLES_DIR, f) for f in os.listdir(SAMPLES_DIR) if f.lower().endswith(('.xml', '.musicxml'))]

# --- Build Gradio Interface ---
with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown(
        """
        # Automatic Music Harmony Generator
        Upload your melody in MusicXML (`.xml`) format, and this tool will generate a harmony line.
        1. Upload a `.xml` or `.musicxml` file, or choose one of the samples below.
        2. Click "Generate Harmony".
        3. The results will appear below: a playable audio file and downloadable `.xml` and `.mid` files.
        """
    )

    with gr.Row():
        with gr.Column(scale=1):
            input_file = gr.File(label="Upload Melody MusicXML File", file_types=[".xml", ".musicxml"])
            if sample_files:
                gr.Examples(examples=sample_files, inputs=input_file, label="Or Select a Sample")
            generate_btn = gr.Button("Generate Harmony", variant="primary")

        with gr.Column(scale=2):
            output_audio = gr.Audio(label="Harmony Audio (.wav)", type="filepath")
            with gr.Row():
                output_xml = gr.File(label="Download Harmony MusicXML (.xml)")
                output_mid = gr.File(label="Download Harmony MIDI (.mid)")
            output_logs = gr.Textbox(
                label="Process Logs",
                lines=10,
                max_lines=20,
                interactive=False,
                autoscroll=True
            )
            
    gr.Markdown("---") # Optional: adds a horizontal line for separation
    gr.Markdown("<p style='text-align:center; font-style:italic; color:grey;'>This work in progress innovation is a student initiated research project and is currently for academic and educational purposes only.</p>")

    generate_btn.click(
        fn=generate_harmony_and_outputs,
        inputs=[input_file],
        outputs=[output_xml, output_mid, output_audio, output_logs]
    )

if __name__ == "__main__":
    # Check for dependencies before launching
    if not os.path.exists(SOUNDFONT_PATH):
        print(f"FATAL ERROR: SoundFont file '{SOUNDFONT_PATH}' not found.")
        print("Please download a SoundFont (.sf2 file) and place it in the same directory as this script.")
    elif not os.path.exists(CHORD_TABLE_FILE):
         print(f"FATAL ERROR: Chord table file '{CHORD_TABLE_FILE}' not found.")
    else:
        print("Launching Gradio App...")
        demo.launch(server_name="0.0.0.0", server_port=7860, share=False)