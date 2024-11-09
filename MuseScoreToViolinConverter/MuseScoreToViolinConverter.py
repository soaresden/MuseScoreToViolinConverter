import xml.etree.ElementTree as ET
from functools import lru_cache
from music21 import converter, note, stream
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import zipfile
import os
import json
import threading
import subprocess
import re
from bs4 import BeautifulSoup
import pyperclip


current_file_path = None

# Function to load a MusicXML, MSCX, or MSCZ file
def load_musicxml():
    global current_file_path
    file_path = filedialog.askopenfilename(filetypes=[("All MusicXML, MSCX, and MSCZ files", "*.musicxml;*.mscx;*.mscz")])
    if file_path:
        current_file_path = file_path
        
        # Check for existing HTML file
        directory = os.path.dirname(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        html_path = os.path.join(directory, f"{base_name}_fingering.html")
        
        # Process the original file
        if file_path.endswith('.mscz'):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                extracted_files = [file for file in zip_ref.namelist() if not file.endswith(('.dat', '.tar'))]
                mscx_files = [f for f in extracted_files if f.endswith('.mscx')]
                if mscx_files:
                    zip_ref.extract(mscx_files[0], '/mnt/data')
                    mscx_file_path = os.path.join('/mnt/data', mscx_files[0])
                    score = parse_mscx(mscx_file_path)
                    if score:
                        display_measures(score, is_mscx=True)
                    else:
                        print("No measures found in the MSCX file.")
        elif file_path.endswith('.mscx'):
            score = parse_mscx(file_path)
            if score:
                display_measures(score, is_mscx=True)
            else:
                print("No measures found in the MSCX file.")
        else:
            score = converter.parse(file_path)
            display_measures(score, is_mscx=False)

        file_label.config(text=f"File opened: {os.path.basename(file_path)}")
        
        # Convert notes to violin fingering
        convert_to_violin()
        
        # Load existing HTML file into textbox3 if it exists
        if os.path.exists(html_path):
            with open(html_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
                # Utiliser BeautifulSoup pour extraire le contenu du corps tout en conservant les balises
                soup = BeautifulSoup(content, 'html.parser')
                body_content = soup.find('body')
                
                if body_content:
                    # Effacer le contenu actuel de editable_notes_display
                    editable_notes_display.delete('1.0', tk.END)
                    
                    # Ajouter le contenu avec formatage
                    for line in body_content.find_all('div', class_='measure'):
                        for element in line:
                            if element.name == 'span':
                                # Récupérer le texte et les classes de style
                                text = element.get_text()
                                classes = element.get('class', [])
                                tag = ' '.join(classes)
                                
                                # Insérer le texte avec le tag de style
                                editable_notes_display.insert(tk.END, text, tag)
                            else:
                                # Insérer le texte sans style
                                editable_notes_display.insert(tk.END, element)
                        # Ajouter un saut de ligne pour chaque mesure
                        editable_notes_display.insert(tk.END, '\n')
        else:
            # Si le fichier HTML n'existe pas, charger le fichier normalement
            score = converter.parse(file_path)
            display_measures(score, is_mscx=False)
        
        # Mettre à jour le label avec le nom du fichier ouvert
        file_label.config(text=f"File opened: {os.path.basename(file_path)}")
    else:
        # Copy content from textbox2 to textbox3 with formatting
        copy_text_with_formatting(converted_notes_display, editable_notes_display)
        # Save initial HTML file
        save_as_html(current_file_path)
 
# Function to parse an MSCX file and extract musical information
def parse_mscx(file_path):
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        measures = []
        for measure in root.findall(".//Measure"):
            measure_notes = []
            for chord in measure.findall(".//Chord"):
                for note_elem in chord.findall(".//Note"):
                    pitch_elem = note_elem.find(".//pitch")
                    if pitch_elem is not None and pitch_elem.text is not None:
                        try:
                            midi_number = int(pitch_elem.text)
                            note_name = midi_to_note_name(midi_number)
                            measure_notes.append(note_name)
                        except ValueError:
                            measure_notes.append("?")
                            measure_notes.append(pitch_elem.text)
            if measure_notes:
                measures.append(measure_notes)
        if not measures:
            print("No notes found in the MSCX file.")
        return measures  # Return measures as a list for an MSCX file
    except ET.ParseError as e:
        print(f"Error parsing MSCX file: {e}")
        return []

# Highlight notes in the two Textboxes
def highlight_selection(event=None):
    try:
        widget = root.focus_get()
        if not isinstance(widget, ScrolledText):
            return

        # Remove previous highlights from all textboxes
        for text_widget in [original_notes_display, converted_notes_display, editable_notes_display]:
            text_widget.tag_remove('highlight', '1.0', tk.END)

        try:
            selection_start = widget.index(tk.SEL_FIRST)
            selection_end = widget.index(tk.SEL_LAST)
        except tk.TclError:
            return

        # Get measure number and selected text
        start_line = int(selection_start.split('.')[0])
        start_col = int(selection_start.split('.')[1])
        end_col = int(selection_end.split('.')[1])
        
        # Get the full line content
        line_content = widget.get(f"{start_line}.0", f"{start_line}.end")
        
        # Count non-space characters up to the selection
        non_space_chars_before = len([c for c in line_content[:start_col] if c != ' '])
        non_space_chars_to_end = len([c for c in line_content[:end_col] if c != ' '])
        
        # Format selection display
        if non_space_chars_before == non_space_chars_to_end - 1:
            selection_display = f"M{start_line:03d}:[{non_space_chars_before}]"
        else:
            selection_display = f"M{start_line:03d}:[{non_space_chars_before}:{non_space_chars_to_end-1}]"
        
        # Pour chaque textbox
        for text_widget in [original_notes_display, converted_notes_display, editable_notes_display]:
            text_widget.tag_add('highlight', selection_start, selection_end)
            text_widget.tag_configure('highlight', background='yellow')

        # Update selection label
        selection_label.config(text=selection_display)

    except Exception as e:
        print(f"Error in highlight_selection: {str(e)}")
        
# Function to open the score in MuseScore
def copy_with_format():
    try:
        widget = root.focus_get()
        if not isinstance(widget, ScrolledText):
            return
            
        selection = widget.get(tk.SEL_FIRST, tk.SEL_LAST)
        tags = widget.tag_names(tk.SEL_FIRST)
        
        root.clipboard_clear()
        root.clipboard_append(selection)  # Simplifié pour ne copier que le texte
    except tk.TclError:
        pass

def paste_with_format():
    try:
        widget = root.focus_get()
        if not isinstance(widget, ScrolledText):
            return
            
        clipboard_data = root.clipboard_get()
        widget.insert(tk.INSERT, clipboard_data)  # Colle directement le texte
    except tk.TclError:
        pass
    
def check_musescore_path():
    musescore_paths = [
        r"C:\Program Files\MuseScore 4\bin\MuseScore4.exe",
        r"C:\Program Files\MuseScore 3\bin\MuseScore3.exe",
        r"C:\Program Files (x86)\MuseScore 4\bin\MuseScore4.exe",
        r"C:\Program Files (x86)\MuseScore 3\bin\MuseScore3.exe"
    ]
    
    for path in musescore_paths:
        if os.path.exists(path):
            return True
    return False

def open_in_musescore():
    global current_file_path
    if current_file_path and os.path.exists(current_file_path):
        try:
            root.config(cursor="wait")  # Change le curseur en sablier
            musescore_button.config(state='disabled')  # Désactive le bouton pendant le chargement
            
            # Met à jour l'interface
            root.update()
            
            # Recherche du chemin MuseScore
            musescore_paths = [
                r"C:\Program Files\MuseScore 4\bin\MuseScore4.exe",
                r"C:\Program Files\MuseScore 3\bin\MuseScore3.exe",
                r"C:\Program Files (x86)\MuseScore 4\bin\MuseScore4.exe",
                r"C:\Program Files (x86)\MuseScore 3\bin\MuseScore3.exe"
            ]
            
            for path in musescore_paths:
                if os.path.exists(path):
                    subprocess.Popen([path, current_file_path])
                    break
                    
        except Exception as e:
            print(f"Error opening MuseScore: {str(e)}")
        finally:
            root.config(cursor="")  # Remet le curseur normal
            musescore_button.config(state='normal')  # Réactive le bouton
    else:
        print("No file currently opened")
        
# Create Dictionnaries for Strings
def get_g_string_fingering():
    return [
        ('0', 'G3', 'G String'),
        ('1₁', 'G#3/Ab3', 'G String'),
        ('1', 'A3', 'G String'),
        ('1¹', 'A#3/Bb3', 'G String'),
        ('2', 'B3', 'G String'),
        ('3₁', 'C4', 'G String'),
        ('3¹', 'C#4/Db4', 'G String'),
        ('4', 'D4', 'G String')
    ]

def get_d_string_fingering():
    return [
        ('0', 'D4', 'D String'),
        ('1₁', 'D#4/Eb4', 'D String'),
        ('1', 'E4', 'D String'),
        ('1', 'E-4', 'D String'),
        ('1¹', 'F4', 'D String'),
        ('2', 'F#4/Gb4', 'D String'),
        ('3', 'G4', 'D String'),
        ('3¹', 'G#4/Ab4', 'D String'),
        ('4', 'A4', 'D String')
    ]
    
def get_a_string_fingering():
    return [
        ('0', 'A4', 'A String'),
        ('1₁', 'A#4/Bb4', 'A String'),
        ('1', 'B4', 'A String'),
        ('1¹', 'C5', 'A String'),
        ('2', 'C#5/Db5', 'A String'),
        ('3₁', 'D5', 'A String'),
        ('3¹', 'D#5/Eb5', 'A String'),
        ('4', 'E5', 'A String'),
        ('4', 'E-5', 'A String')
    ]

# 1st Position
def get_e_string_fingering():
    return [
        ('0', 'E5', 'E String'),
        ('0', 'E-5', 'E String'),
        ('1₁', 'F5', 'E String'),
        ('1', 'F#5/Gb5', 'E String'),
        ('1¹', 'G5', 'E String'),
        ('2', 'G#5/Ab5', 'E String'),
        ('3₁', 'A5', 'E String'),
        ('3¹', 'A#5/Bb5', 'E String'),
        ('4', 'B5', 'E String')
    ]

# 2nd position
def get_e_string_fingering_position2():
    return [
        ('1', 'F5', 'E String'),
        ('1¹', 'F#5/Gb5', 'E String'),
        ('2', 'G5', 'E String'),
        ('2¹', 'G#5/Ab5', 'E String'),
        ('3', 'A5', 'E String'),
        ('3¹', 'A#5/Bb5', 'E String'),
        ('4', 'B5', 'E String'),
        ('4¹', 'C6', 'E String')
    ]

# 3th position
def get_e_string_fingering_position3():
    return [
        ('1', 'G5', 'E String'),
        ('1¹', 'G#5/Ab5', 'E String'),
        ('2', 'A5', 'E String'),
        ('2¹', 'A#5/Bb5', 'E String'),
        ('3', 'B5', 'E String'),
        ('3¹', 'C6', 'E String'),
        ('4', 'C#6/Db6', 'E String'),
        ('4¹', 'D6', 'E String')
    ]

# 4th position
def get_e_string_fingering_position4():
    return [
        ('1', 'A5', 'E String'),
        ('1¹', 'A#5/Bb5', 'E String'),
        ('2', 'B5', 'E String'),
        ('2¹', 'C6', 'E String'),
        ('3', 'C#6/Db6', 'E String'),
        ('3¹', 'D6', 'E String'),
        ('4', 'D#6/Eb6', 'E String'),
        ('4¹', 'E6', 'E String')
    ]

# 5th position
def get_e_string_fingering_position5():
    return [
        ('1', 'B5', 'E String'),
        ('1¹', 'C6', 'E String'),
        ('2', 'C#6/Db6', 'E String'),
        ('2¹', 'D6', 'E String'),
        ('3', 'D#6/Eb6', 'E String'),
        ('3¹', 'E6', 'E String'),
        ('4', 'F6', 'E String'),
        ('4¹', 'F#6/Gb6', 'E String')
    ]

# 6th position
def get_e_string_fingering_position6():
    return [
        ('1', 'C6', 'E String'),
        ('1¹', 'C#6/Db6', 'E String'),
        ('2', 'D6', 'E String'),
        ('2¹', 'D#6/Eb6', 'E String'),
        ('3', 'E6', 'E String'),
        ('3¹', 'F6', 'E String'),
        ('4', 'F#6/Gb6', 'E String'),
        ('4¹', 'G6', 'E String')
    ]

# 7th position
def get_e_string_fingering_position7():
    return [
        ('1', 'D6', 'E String'),
        ('1¹', 'D#6/Eb6', 'E String'),
        ('2', 'E6', 'E String'),
        ('2¹', 'F6', 'E String'),
        ('3', 'F#6/Gb6', 'E String'),
        ('3¹', 'G6', 'E String'),
        ('4', 'G#6/Ab6', 'E String'),
        ('4¹', 'A6', 'E String')
    ]

# Associate colors string (ColourStrings Method)
string_colors = {
    'G String': 'green',   # for G
    'D String': 'red',     # Red for D
    'A String': 'blue',    # Blue for A
    'E String': 'brown'    # Brown instead of Yellow because can't see nothing for E
}
# Update Positions Colors
position_colors = {
    '1': {'background': 'white', 'foreground': 'black'},
    '2': {'background': 'gray', 'foreground': 'white'},
    '3': {'background': 'brown', 'foreground': 'white'},
    '4': {'background': 'purple', 'foreground': 'white'},
    '5': {'background': 'pink', 'foreground': 'white'},
    '6': {'background': 'turquoise', 'foreground': 'black'},
    '7': {'background': 'blue', 'foreground': 'white'},
}

# Set of notes below the violin range
low_notes = {
    'A0', 'A#0', 'Bb0', 'B0', 'C1', 'C#1', 'Db1', 'D1', 'D#1', 'Eb1', 'E1'
}

# Set of notes above the violin range
high_notes = {
    'F6', 'F#6', 'Gb6', 'G6', 'G#6', 'Ab6', 'A6', 'A#6', 'Bb6', 'B6', 'C7', 'C#7',
    'Db7', 'D7', 'D#7', 'Eb7', 'E7', 'F7', 'F#7', 'Gb7', 'G7'
}

# Function to display musical measures in different text areas
def display_measures(score, is_mscx=False):
    # Clear previous content
    original_notes_display.delete('1.0', tk.END)
    converted_notes_display.delete('1.0', tk.END)
    editable_notes_display.delete('1.0', tk.END)

    # If the score is a list (case of an MSCX file)
    if is_mscx:
        measures = score
    else:
        # Extract notes per measure, considering repeats (MusicXML case)
        measures = []
        for part in score.parts:
            for measure in part.getElementsByClass(stream.Measure):
                measure_notes = []
                for element in measure.flatten().notes:  # Use .flatten() instead of .flat
                    if isinstance(element, note.Note):
                        measure_notes.append(element)
                measures.append(measure_notes)

    # Display notes per measure with string color
    for i, measure in enumerate(measures):
        block_notes_str = " ".join([f"{n.nameWithOctave}" if isinstance(n, note.Note) else n for n in measure])
        measure_label = f"M{i+1:03d}: "
        original_notes_display.insert(tk.END, f"{measure_label}{block_notes_str}\n")  # Add newline

# Function to convert notes to violin fingering
def convert_to_violin():
    lines = original_notes_display.get('1.0', tk.END).splitlines()
    for i, line in enumerate(lines):
        if line.startswith("M"):
            notes_str = line[5:]
            notes = notes_str.split(" ")
            measure_label = f"M{i+1:03d}: "
            converted_notes_display.insert(tk.END, f"{measure_label}")
            editable_notes_display.insert(tk.END, f"{measure_label}")
            for note_name in notes:
                if note_name == "":
                    continue

                # Check if the note is outside the violin range
                if note_name in low_notes:
                    # Notes lower than G3
                    converted_notes_display.insert(tk.END, f"{note_name}", (f"low_{i}_{note_name}",))
                    editable_notes_display.insert(tk.END, f"{note_name}", (f"low_{i}_{note_name}",))
                    converted_notes_display.tag_config(f"low_{i}_{note_name}", foreground='green', background='black')
                    editable_notes_display.tag_config(f"low_{i}_{note_name}", foreground='green', background='black')
                    converted_notes_display.insert(tk.END, " ")
                    editable_notes_display.insert(tk.END, " ")
                    continue
                elif note_name in high_notes:
                    # Notes higher than E6
                    converted_notes_display.insert(tk.END, f"{note_name}", (f"high_{i}_{note_name}",))
                    editable_notes_display.insert(tk.END, f"{note_name}", (f"high_{i}_{note_name}",))
                    converted_notes_display.tag_config(f"high_{i}_{note_name}", foreground='turquoise', background='black')
                    editable_notes_display.tag_config(f"high_{i}_{note_name}", foreground='turquoise', background='black')
                    converted_notes_display.insert(tk.END, " ")
                    editable_notes_display.insert(tk.END, " ")
                    continue

                # Then Check G, D, and A strings
                possibilities = []
                for df in [fingering_df_g_string, fingering_df_d_string, fingering_df_a_string]:
                    result = df[df['Note'].str.contains(note_name, na=False)]
                    if not result.empty:
                        possibilities.append(result.iloc[0])

                if len(possibilities) == 1:
                    finger = possibilities[0]['Fingers Used']
                    string = possibilities[0]['String']
                elif len(possibilities) > 1:
                    best_choice = min(possibilities, key=lambda x: int(x['Fingers Used'].strip('₁¹')))
                    finger = best_choice['Fingers Used']
                    string = best_choice['String']
                else:
                    finger, string = "?", None

                color = string_colors.get(string, 'black')
                if finger != "?":
                    converted_notes_display.insert(tk.END, f"{finger}", (f"color_{i}_{note_name}",))
                    editable_notes_display.insert(tk.END, f"{finger}", (f"color_{i}_{note_name}",))
                    converted_notes_display.tag_config(f"color_{i}_{note_name}", foreground=color)
                    editable_notes_display.tag_config(f"color_{i}_{note_name}", foreground=color)
                    converted_notes_display.insert(tk.END, " ")
                    editable_notes_display.insert(tk.END, " ")
                    continue

                # Try successive positions on the E string (priority order)
                possibilities = []
                for position, get_fingering_function in enumerate(
                    [get_e_string_fingering, get_e_string_fingering_position3, get_e_string_fingering_position5,
                     get_e_string_fingering_position7, get_e_string_fingering_position2, get_e_string_fingering_position4,
                     get_e_string_fingering_position6], start=1):
                    
                    possibilities += [pos for pos in get_fingering_function() if note_name in pos[1]]

                if possibilities:
                    best_choice = min(possibilities, key=lambda x: int(x[0].strip('₁¹')))
                    finger = best_choice[0]
                    string = best_choice[2]
                    color_info = position_colors[str(position)]
                    converted_notes_display.insert(tk.END, f"{finger}", (f"position_{i}_{note_name}",))
                    editable_notes_display.insert(tk.END, f"{finger}", (f"position_{i}_{note_name}",))
                    converted_notes_display.tag_config(f"position_{i}_{note_name}",
                                                       foreground=color_info['foreground'],
                                                       background=color_info['background'])
                    editable_notes_display.tag_config(f"position_{i}_{note_name}",
                                                      foreground=color_info['foreground'],
                                                      background=color_info['background'])
                    converted_notes_display.insert(tk.END, " ")
                    editable_notes_display.insert(tk.END, " ")
                    continue

                # If the note does not match any category
                converted_notes_display.insert(tk.END, f"?", (f"unknown_{i}_{note_name}",))
                editable_notes_display.insert(tk.END, f"?", (f"unknown_{i}_{note_name}",))
                converted_notes_display.tag_config(f"unknown_{i}_{note_name}", foreground='white', background='red')
                editable_notes_display.tag_config(f"unknown_{i}_{note_name}", foreground='white', background='red')
                converted_notes_display.insert(tk.END, " ")
                editable_notes_display.insert(tk.END, " ")

            converted_notes_display.insert(tk.END, "\n")
            editable_notes_display.insert(tk.END, "\n")

# Save modifications when the content of editable_notes_display changes
# Modifiez la fonction `on_edit` pour détecter l'événement "Entrée"
def on_edit(event):
    global current_file_path
    if event.keysym == 'Return':  # Vérifie si la touche pressée est "Entrée"
        if current_file_path:
            save_as_html(current_file_path)  # Sauvegarde du fichier HTML
            file_label.config(text=f"Last saved: {os.path.basename(current_file_path)}")
        editable_notes_display.edit_modified(False)


# Function to convert MIDI number to note name
def midi_to_note_name(midi_number):
    note_names = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
    octave = (midi_number // 12) - 1
    note_index = midi_number % 12
    return f"{note_names[note_index]}{octave}"

# Change the color of the selection in the third Textbox
def change_color(color):
    if color:
        try:
            selection_start = editable_notes_display.index(tk.SEL_FIRST)
            selection_end = editable_notes_display.index(tk.SEL_LAST)
        except tk.TclError:
            return
        tag_name = f'color_{selection_start}_{selection_end}'
        editable_notes_display.tag_add(tag_name, selection_start, selection_end)
        editable_notes_display.tag_configure(tag_name, foreground=color)
        
# Replace save_modifications function with save_as_html
def save_as_html(file_path):
    if not file_path:
        return
    # Définir le chemin du fichier HTML
    directory = os.path.dirname(file_path)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    html_path = os.path.join(directory, f"{base_name}_fingering.html")

    # Préparer le contenu HTML avec les styles
    html_content = ['<!DOCTYPE html>', '<html>', '<head>', '<style>']
    html_content.append('''
        body { white-space: pre-wrap; font-family: monospace; }
        .default { color: black; }  /* Style par défaut */
    ''')

    # Récupérer les styles de chaque tag utilisé dans editable_notes_display
    tags_used = {}
    for tag in editable_notes_display.tag_names():
        if tag not in ('sel', 'highlight'):
            config = editable_notes_display.tag_configure(tag)
            style = ""
            if 'foreground' in config:
                style += f"color: {config['foreground'][4]};"
            if 'background' in config:
                style += f" background-color: {config['background'][4]};"
            tags_used[tag] = style

    # Ajouter les styles des tags dans le HTML
    for tag, style in tags_used.items():
        html_content.append(f'.{tag} {{ {style} }}')
    html_content.extend(['</style>', '</head>', '<body>'])

    # Traitement ligne par ligne du contenu de editable_notes_display
    total_lines = int(editable_notes_display.index("end-1c").split(".")[0])
    for line_num in range(1, total_lines + 1):
        line_start = f"{line_num}.0"
        line_end = f"{line_num}.end"
        line_text = editable_notes_display.get(line_start, line_end)

        # Démarrer une nouvelle ligne avec un div pour chaque mesure
        html_content.append('<div class="measure">')
        for pos in range(len(line_text)):
            current_pos = f"{line_num}.{pos}"
            char = editable_notes_display.get(current_pos)
            tags = editable_notes_display.tag_names(current_pos)
            
            # Vérification et ajout des tags manquants
            if not tags:
                tags = ['default']  # Utiliser un tag 'default' pour les caractères non stylés

            tag_classes = ' '.join(tags)

            # Encapsuler chaque caractère dans un span avec les classes de tag
            if tag_classes:
                html_content.append(f'<span class="{tag_classes}">{char}</span>')
            else:
                html_content.append(char)  # Ajouter le caractère directement si pas de tag

        # Fermer le div de la mesure avec un saut de ligne visuel
        html_content.append('</div>')

    html_content.extend(['</body>', '</html>'])

    # Écrire le contenu HTML dans le fichier
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(''.join(html_content))  # Écrire sans ajouter de sauts de ligne

    # Mise à jour de l'interface pour indiquer la sauvegarde
    file_label.config(text=f"Last saved: {os.path.basename(html_path)}")





def refresh_text_display():
    """Refresh the editable_notes_display content to reapply styles and ensure format consistency."""
    content = editable_notes_display.get("1.0", tk.END)
    editable_notes_display.delete("1.0", tk.END)
    editable_notes_display.insert("1.0", content)

    # Reapply all tags to ensure they display correctly after reloading
    for line_num in range(1, int(editable_notes_display.index("end-1c").split(".")[0]) + 1):
        line_start = f"{line_num}.0"
        line_end = f"{line_num}.end"
        line_text = editable_notes_display.get(line_start, line_end)
        for pos in range(len(line_text)):
            current_pos = f"{line_num}.{pos}"
            tags = editable_notes_display.tag_names(current_pos)
            for tag in tags:
                editable_notes_display.tag_add(tag, current_pos, f"{current_pos}+1c")

    



# Function to copy text with formatting from one textbox to another
def copy_with_format():
    try:
        # Récupérer la sélection dans editable_notes_display
        selection_start = editable_notes_display.index(tk.SEL_FIRST)
        selection_end = editable_notes_display.index(tk.SEL_LAST)
        selected_text = editable_notes_display.get(selection_start, selection_end)
        
        # Construire le HTML du texte sélectionné
        html_content = ['<!DOCTYPE html>', '<html>', '<body>', '<pre>']  # Utilise <pre> pour conserver les espaces
        
        current_index = selection_start
        while current_index != selection_end:
            # Récupérer le caractère et ses tags
            char = editable_notes_display.get(current_index)
            tags = editable_notes_display.tag_names(current_index)
            
            # Construire un span avec les styles
            if tags:
                tag_classes = ' '.join(tags)
                html_content.append(f'<span class="{tag_classes}">{char}</span>')
            else:
                html_content.append(char)
                
            # Passer au caractère suivant
            current_index = editable_notes_display.index(f"{current_index}+1c")
        
        html_content.extend(['</pre>', '</body>', '</html>'])
        
        # Convertir en une seule chaîne de HTML
        html_to_copy = ''.join(html_content)

        # Copier le texte brut et le HTML dans le presse-papiers
        pyperclip.copy(selected_text)  # Copie le texte brut
        pyperclip.copy(html_to_copy)  # Copie le HTML formaté

        messagebox.showinfo("Copie", "Texte copié avec formatage HTML dans le presse-papiers!")
    except tk.TclError:
        messagebox.showwarning("Erreur", "Aucune sélection n'a été faite.")



# Creating DataFrames
fingering_df_g_string = pd.DataFrame(get_g_string_fingering(), columns=['Fingers Used', 'Note', 'String'])
fingering_df_d_string = pd.DataFrame(get_d_string_fingering(), columns=['Fingers Used', 'Note', 'String'])
fingering_df_a_string = pd.DataFrame(get_a_string_fingering(), columns=['Fingers Used', 'Note', 'String'])
fingering_df_e_string = pd.DataFrame(get_e_string_fingering(), columns=['Fingers Used', 'Note', 'String'])



# Set up GUI components
root = tk.Tk()
root.title("MusicXML Note & Fingering Display")
root.geometry("1200x600")  # Set window size
root.bind('<Control-c>', lambda e: copy_with_format())
root.bind('<Control-v>', lambda e: paste_with_format())

# Button to open a MusicXML, MSCX, or MSCZ file
open_button = tk.Button(root, text="Open a MusicXML/MSCX/MSCZ File", command=load_musicxml)
open_button.pack(pady=10)

# Label to display the opened file name
file_label = tk.Label(root, text="No file opened")
file_label.pack(pady=5)

# Ajouter un label pour la sélection des notes
selection_label = tk.Label(root, padx=0, pady=10, text="Selected Notes: None")
selection_label.place(x=0, y=50)

# Buttons to Apply Color to Text
green_button = tk.Button(root, text=" ", bg="green", command=lambda: change_color("green"))
green_button.place(x=850, y=50, width=25)

red_button = tk.Button(root, text=" ", bg="red", command=lambda: change_color("red"))
red_button.place(x=875, y=50, width=25)

blue_button = tk.Button(root, text=" ", bg="blue", command=lambda: change_color("blue"))
blue_button.place(x=900, y=50, width=25)

brown_button = tk.Button(root, text=" ", bg="brown", command=lambda: change_color("brown"))
brown_button.place(x=925, y=50, width=25)


# Add Legend for E String
legend_label = tk.Label(root, text="E String:")
legend_label.place(x=900, y=20)
# Apply legend
legend_colors = ['white', 'gray', 'brown', 'purple', 'pink', 'turquoise', 'blue']
for i, color in enumerate(legend_colors, start=1):
    legend_label = tk.Label(root, text=f"[{i}]", bg=color, fg='black' if color == 'white' else ('white' if color != 'turquoise' else 'black'))
    legend_label.place(x=945 + (i * 30), y=20)
   

# Open in MuseScore
musescore_button = tk.Button(root, text="Open in MuseScore", command=open_in_musescore)
if check_musescore_path():
    musescore_button.place(x=1010, y=50)

# ScrolledText for original notes display
original_notes_display = ScrolledText(root, width=50, height=10)
original_notes_display.pack(pady=10, side=tk.LEFT, fill=tk.BOTH, expand=True)
original_notes_display.insert(tk.END, "Original Notes:\n")
original_notes_display.bind('<<Selection>>', highlight_selection)

# ScrolledText for converted notes display
converted_notes_display = ScrolledText(root, width=50, height=10)
converted_notes_display.pack(pady=10, side=tk.LEFT, fill=tk.BOTH, expand=True)
converted_notes_display.insert(tk.END, "Converted Violin Notes:\n")
converted_notes_display.bind('<<Selection>>', highlight_selection)

# ScrolledText for editable notes display
editable_notes_display = ScrolledText(root, width=50, height=10)
editable_notes_display.pack(pady=10, side=tk.RIGHT, fill=tk.BOTH, expand=True)
editable_notes_display.insert(tk.END, "Editable Notes:\n")
editable_notes_display.bind('<<Selection>>', highlight_selection)
editable_notes_display.bind('<Return>', on_edit)
root.bind('<Control-c>', lambda e: copy_with_format())


root.mainloop()
