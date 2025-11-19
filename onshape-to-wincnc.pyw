"""
GUI Application for Converting Onshape CAM Studio G-code to WinCNC Format
=======================================================================

This module provides a simple graphical user interface (GUI) that allows
operators to select a G-code file exported from Onshape CAM Studio and
convert it to a WinCNC-compatible format for ShopSabre routers.  The
interface guides the user through selecting the input file, displays the
derived output file name (prefixed with ``SS23_``), and provides feedback
during the conversion process.  A confirmation dialog alerts the user
when the conversion succeeds or if an error occurs.

* Converts parentheses comments to WinCNC-friendly square brackets and
  removes semicolon comments.
* Splits spindle speed (``S``) and spindle start/stop commands
  (``M3``/``M4``/``M5``) onto separate lines.
* Inserts the last arc command (``G2``/``G3``) on subsequent lines
  containing arc parameters because arcs are not modal in WinCNC.
* Optionally removes tool changes (``M6``) for machines without
  automatic tool changers.
* Optionally removes mist coolant codes (``M7`` / ``M9``) if they are
  not relevant for the machine, or rewrites them to WinCNC's
  ``M11C<port>`` / ``M12C<port>`` mister control when configured.
* Checks for likely Z-zero / Position Type mistakes, with behavior
  dependent on the selected Zero Plane ("Top" or "Bottom").
"""

import json
import os
import re
from dataclasses import dataclass, asdict
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional


SETTINGS_FILE = Path.home() / '.onshape_to_wincnc_settings.json'


def _env_int(name: str, default: Optional[int]) -> Optional[int]:
    """Helper to safely parse integer environment overrides."""

    raw = os.environ.get(name)
    if raw is None:
        return default
    raw = raw.strip()
    if not raw:
        return default
    try:
        value = int(raw, 10)
    except ValueError:
        return default
    return value if value > 0 else None


@dataclass
class MachineSettings:
    """Represents user-editable ShopSabre integration parameters."""

    mist_port: Optional[int] = None
    output_directory: Optional[str] = None
    output_name_mode: str = 'prefix'
    output_name_value: str = 'SS23_'

    @staticmethod
    def _coerce_channel(value, fallback: Optional[int]) -> Optional[int]:
        if value in (None, ''):
            return None
        try:
            parsed = int(value)
        except (TypeError, ValueError):
            return fallback
        return parsed if parsed > 0 else fallback

    @classmethod
    def load(cls) -> 'MachineSettings':
        defaults = {
            'mist_port': _env_int('SHOP_SABRE_MIST_PORT', _env_int('SHOP_SABRE_MIST_OUTPUT', None)),
            'output_directory': None,
            'output_name_mode': 'prefix',
            'output_name_value': 'SS23_',
        }
        data = {}
        if SETTINGS_FILE.exists():
            try:
                data = json.loads(SETTINGS_FILE.read_text(encoding='utf-8'))
            except Exception:
                data = {}
        mist = cls._coerce_channel(
            data.get('mist_port', data.get('mist_output')),
            defaults['mist_port']
        )
        raw_directory = data.get('output_directory')
        directory = str(raw_directory).strip() if raw_directory else ''
        directory = os.path.abspath(os.path.expanduser(directory)) if directory else None
        mode_raw = str(data.get('output_name_mode', defaults['output_name_mode'])).strip().lower()
        mode = mode_raw if mode_raw in ('prefix', 'suffix') else defaults['output_name_mode']
        name_value = str(data.get('output_name_value', defaults['output_name_value']))
        return cls(
            mist_port=mist,
            output_directory=directory,
            output_name_mode=mode,
            output_name_value=name_value,
        )

    def save(self) -> None:
        SETTINGS_FILE.write_text(json.dumps(asdict(self), indent=2), encoding='utf-8')


SETTINGS = MachineSettings.load()

def parentheses_to_bracket_lines(line: str) -> tuple[str, list[str]]:
    """Convert inline parentheses comments into bracket-only lines.

    The returned content line excludes parentheses segments. Each
    extracted comment is converted to a standalone ``[comment]`` line to
    satisfy WinCNC's preference for bracketed comments on their own
    lines. Nested parentheses are supported; unmatched parentheses leave
    the remainder untouched to avoid dropping user data.
    """

    content_chars = []
    bracket_comments: list[str] = []
    i = 0
    length = len(line)
    while i < length:
        char = line[i]
        if char == '(':
            i += 1
            depth = 1
            comment_start = i
            while i < length and depth > 0:
                if line[i] == '(':
                    depth += 1
                elif line[i] == ')':
                    depth -= 1
                    if depth == 0:
                        comment = line[comment_start:i].strip()
                        if comment:
                            bracket_comments.append(f'[{comment}]')
                        break
                i += 1
            if depth == 0:
                i += 1  # Skip the closing parenthesis
            else:
                # Unmatched parentheses; keep the raw remainder
                content_chars.append('(')
                content_chars.append(line[comment_start:])
                break
        else:
            content_chars.append(char)
            i += 1
    content_line = ''.join(content_chars).strip()
    return content_line, bracket_comments


def remove_semicolon_comments(line: str) -> str:
    """Remove everything after a semicolon comment."""
    if ';' in line:
        return line.split(';', 1)[0]
    return line


def split_spindle_speed_and_m(line: str):
    """Split S-word from M3/M4/M5 if they appear together in one block."""
    tokens = line.strip().split()
    s_index = None
    m_index = None
    for i, token in enumerate(tokens):
        if token.upper().startswith('S') and token[1:].replace('.', '').isdigit():
            s_index = i
        if token.upper() in ('M3', 'M4', 'M5'):
            m_index = i
    if s_index is not None and m_index is not None and m_index > s_index:
        return [' '.join(tokens[:m_index]), ' '.join(tokens[m_index:])]
    return [line]


def process_arc_line(line: str, last_g: str):
    """Maintain non-modal behaviour for G2/G3.

    If a line contains no G-code but has arc centre parameters (I/J/K/R)
    following a G2/G3, prefix the previous G-code.
    """
    stripped = line.strip()
    if not stripped:
        return line, last_g
    tokens = stripped.split()
    gcode_pattern = re.compile(r'G0?\d+(\.\d+)?', re.IGNORECASE)
    g_found = None
    for t in tokens:
        if gcode_pattern.match(t):
            g_found = t.upper()
            break
    if g_found:
        last_g = g_found
        return line, last_g
    # If the last G was G2/G3 and arc params present, prefix it
    if last_g in ('G2', 'G02', 'G3', 'G03'):
        if any(tok[0].upper() in ('I', 'J', 'K', 'R') for tok in tokens):
            # Sequence numbers (N words) are removed in the improved converter,
            # so we can simply prefix without worrying about them.
            return f"{last_g} {line.lstrip()}", last_g
    return line, last_g


def remove_unsupported_tokens(tokens, remove_coolant: bool, remove_toolchange: bool, mist_port: Optional[int]):
    """Filter out tokens that WinCNC does not support or should be handled separately.

    Removes program delimiters (%, O#####), line numbers (N####), tool
    length offsets (H#). Optionally removes coolant codes (M7/M8/M9) and
    tool change commands (M6). Omits plane selection, cutter
    compensation and canned cycle cancel codes (G17, G40, G80). G49
    (tool length cancel) is retained and handled in a post-processing
    step. Tool selection (T#) words are preserved when tool changes are
    kept so they can be paired with M6 in WinCNC's expected format.
    """
    result = []
    for tok in tokens:
        up = tok.upper()
        # Skip program delimiters and program numbers
        if up.startswith('%'):
            continue
        if up.startswith('O') and up[1:].isdigit():
            continue
        # Remove line numbers
        if up.startswith('N') and up[1:].isdigit():
            continue
        # Remove tool selection and length offset words
        if up.startswith('T') and up[1:].replace('.', '').isdigit():
            if remove_toolchange:
                continue
            result.append(up)
            continue
        if up.startswith('H') and up[1:].replace('.', '').isdigit():
            continue
        # Handle coolant commands (mist only)
        if up in ('M7', 'M9'):
            if remove_coolant:
                continue
            if mist_port is None:
                raise ValueError('Mist commands require a configured mister port.')
            converted = f"M11C{mist_port}" if up == 'M7' else f"M12C{mist_port}"
            result.append(converted)
            continue
        # Flood (M8) is not supported; drop it.
        if up == 'M8':
            continue
        # Optionally remove tool change commands
        if remove_toolchange and up == 'M6':
            continue
        # Omit plane selection and cutter compensation codes; these can
        # cause syntax or multiple command errors if combined with other G-codes.
        if up in ('G17', 'G40', 'G80'):
            continue
        result.append(tok)
    return result


def normalize_tool_change(tokens: list[str]) -> list[str]:
    """Ensure tool changes follow the "TX M6" format expected by WinCNC."""

    if not tokens:
        return tokens

    tool_token = None
    has_m6 = False
    for tok in tokens:
        up = tok.upper()
        if up.startswith('T') and up[1:].replace('.', '').isdigit():
            if tool_token is None:
                tool_token = up
            continue
        if up == 'M6':
            has_m6 = True

    if not has_m6:
        return tokens

    remaining = []
    for tok in tokens:
        up = tok.upper()
        if up == 'M6':
            continue
        if up.startswith('T') and up[1:].replace('.', '').isdigit():
            continue
        remaining.append(tok)

    ordered = []
    if tool_token:
        ordered.append(tool_token)
    ordered.append('M6')
    ordered.extend(remaining)
    return ordered


def split_by_multiple_commands(tokens):
    """Divide tokens into groups with at most one G or M command."""
    lines = []
    current = []
    has_command = False
    for tok in tokens:
        if not tok:
            continue
        up = tok.upper()
        is_command = up.startswith('G') or up.startswith('M')
        if is_command:
            if has_command and current:
                lines.append(current)
                current = []
                has_command = False
            current.append(tok)
            has_command = True
        else:
            current.append(tok)
    if current:
        lines.append(current)
    return lines


def get_g_code(token: str):
    m = re.match(r'^G\d+(\.\d+)?', token, re.IGNORECASE)
    if m:
        return m.group(0).upper()
    return None


def convert_lines(
    lines,
    remove_coolant: bool = True,
    remove_toolchange: bool = True,
    mist_port: Optional[int] = None
):
    """Convert a list of G-code lines into WinCNC-friendly format.

    Parameters
    ----------
    lines : list[str]
        Raw G-code lines.
    remove_coolant : bool
        If True, remove M7/M9 commands.
    remove_toolchange : bool
        If True, remove M6 commands.
    mist_port : Optional[int]
        Mister port number used to translate M7/M9 into M11C/M12C when
        coolant commands are kept.
    """
    converted = []
    last_motion = None
    for original_line in lines:
        line = original_line.rstrip('\n')
        line = remove_semicolon_comments(line)
        content_line, bracket_comments = parentheses_to_bracket_lines(line)
        if not content_line.strip() and not bracket_comments:
            converted.append('')
            continue
        for sm_line in split_spindle_speed_and_m(content_line.strip()):
            tokens = sm_line.split()
            tokens = remove_unsupported_tokens(tokens, remove_coolant, remove_toolchange, mist_port)
            if not tokens:
                continue
            tokens = normalize_tool_change(tokens)
            grouped = split_by_multiple_commands(tokens)
            for group in grouped:
                if not group:
                    continue
                g_code_in_line = None
                for tok in group:
                    g_code_in_line = get_g_code(tok)
                    if g_code_in_line:
                        break
                if g_code_in_line:
                    m = re.match(r'^G0?([0123])', g_code_in_line)
                    if m:
                        last_motion = f"G{m.group(1)}"
                else:
                    # If no G-code in this group, prefix last motion on coordinate lines
                    if last_motion:
                        has_arc = any(tok[0].upper() in ('I', 'J', 'K', 'R') for tok in group)
                        has_lin = any(tok[0].upper() in ('X', 'Y', 'Z', 'W') for tok in group)
                        if last_motion in ('G2', 'G3') and has_arc:
                            group.insert(0, last_motion)
                        elif last_motion in ('G0', 'G1') and has_lin:
                            group.insert(0, last_motion)
                converted.append(' '.join(group))
        converted.extend(bracket_comments)
    # After processing all lines, perform post-processing on the converted list.
    # 1. Handle placement of G49 (tool length cancel) commands:
    #    - If there is a spindle stop (M5), remove any G49 before the first M5
    #      and keep those that follow.
    #    - If there is no M5, keep only the last G49 and drop earlier occurrences.
    m5_index = None
    for idx, ln in enumerate(converted):
        if ln.strip().upper().startswith('M5'):
            m5_index = idx
            break
    final_lines = []
    if m5_index is None:
        # Retain only the last G49
        g49_indices = [i for i, ln in enumerate(converted) if ln.strip().upper().startswith('G49')]
        remove_set = set(g49_indices[:-1])
        for i, ln in enumerate(converted):
            if i in remove_set and ln.strip().upper().startswith('G49'):
                continue
            final_lines.append(ln)
    else:
        seen_m5 = False
        for ln in converted:
            stripped = ln.strip().upper()
            if stripped.startswith('M5'):
                seen_m5 = True
                final_lines.append(ln)
                continue
            if stripped.startswith('G49'):
                if not seen_m5:
                    continue
                else:
                    final_lines.append(ln)
                    continue
            final_lines.append(ln)
    # 2. Ensure there are two blank lines following the first G90 near the top
    for idx, line in enumerate(final_lines):
        if line.strip().upper() == 'G90':
            blank_count = 0
            j = idx + 1
            while j < len(final_lines) and final_lines[j].strip() == '':
                blank_count += 1
                j += 1
            while blank_count < 2:
                final_lines.insert(idx + 1, '')
                blank_count += 1
            break
    return final_lines


def detect_z_zero_issue(lines, zero_plane: str = "Top") -> bool:
    """Detect whether the CAM output appears to have an incorrect Z-zero.

    Behavior depends on the selected zero_plane:

    * "Top": Assumes Z=0 is the top of the stock. After spindle on (M3/M4),
      we expect at least one negative Z (cutting into the material). If no
      Z- is ever seen, we assume a likely Position Type / Z-zero mistake.

    * "Bottom": Assumes Z=0 is the bottom of the part. After spindle on
      we expect Z to stay at or above 0 (Z >= 0). If any negative Z is
      seen, we flag it as an issue (cutting below the part bottom / table).

    * "Ignore": Disables Z-zero detection entirely.

    Args:
        lines: List of strings representing the input G-code file.
        zero_plane: "Top" or "Bottom".

    Returns:
        True if a potential Z-zero problem is detected, False otherwise.
    """
    zero_plane = (zero_plane or "Top").strip().title()

    if zero_plane == "Ignore":
        return False
    spindle_on = False
    saw_negative_z = False

    for line in lines:
        uline = line.upper()
        # Detect spindle on commands
        if not spindle_on and ('M3' in uline or 'M4' in uline):
            spindle_on = True
            continue
        if not spindle_on:
            continue

        # Normalize away spaces so 'Z -0.050' still matches
        compact = uline.replace(' ', '')
        if 'Z-' in compact:
            saw_negative_z = True
            # For bottom-zero we can early-out: one negative Z is enough to flag.
            if zero_plane == "Bottom":
                return True

    if zero_plane == "Top":
        # Top-zero heuristic: if we never saw a negative Z after spindle on, warn.
        if not saw_negative_z:
            return True
        return False
    else:
        # Bottom-zero heuristic: if we got here, we haven't seen any Z-, so we're OK.
        return False


def convert_file(
    input_path: str,
    output_path: str,
    remove_coolant: bool = True,
    remove_toolchange: bool = True,
    mist_port: Optional[int] = None
) -> None:
    """Convert the input G-code file and write to the given output path.

    Parameters
    ----------
    input_path : str
        Input G-code file path.
    output_path : str
        Output G-code file path.
    remove_coolant : bool
        If True, remove M7/M9 commands.
    remove_toolchange : bool
        If True, remove M6 commands.
    mist_port : Optional[int]
        Mister port number used to translate M7/M9 into M11C/M12C when
        coolant commands are kept.

    Raises
    ------
    IOError
        On file errors.
    """
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f_in:
        lines = f_in.readlines()
    converted = convert_lines(
        lines,
        remove_coolant=remove_coolant,
        remove_toolchange=remove_toolchange,
        mist_port=mist_port
    )
    with open(output_path, 'w', encoding='utf-8') as f_out:
        for cl in converted:
            f_out.write(cl.rstrip() + '\n')


class ConverterGUI:
    """Graphical interface for Onshape to WinCNC conversion."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title('Onshape to WinCNC Converter')
        self.root.configure(bg='#f4f6fb')
        self.settings = SETTINGS
        self.settings_window: Optional[tk.Toplevel] = None
        self.output_settings_window: Optional[tk.Toplevel] = None

        style = ttk.Style()
        try:
            style.theme_use('clam')
        except tk.TclError:
            # Fallback to default theme if clam is unavailable.
            pass
        style.configure('TFrame', background='#f4f6fb')
        style.configure('Card.TFrame', background='#ffffff', borderwidth=1, relief='solid')
        style.configure('Card.TLabel', background='#ffffff', font=('Segoe UI', 10))
        style.configure('Heading.TLabel', background='#ffffff', font=('Segoe UI Semibold', 11))
        style.configure('Body.TLabel', background='#f4f6fb', font=('Segoe UI', 10))
        style.configure('Status.TLabel', background='#f4f6fb', foreground='#2563eb', font=('Segoe UI', 10))
        style.configure('Accent.TButton', font=('Segoe UI Semibold', 11), padding=(12, 6), foreground='#ffffff',
                        background='#2563eb')
        style.map('Accent.TButton', background=[('active', '#1d4ed8')], foreground=[('disabled', '#d1d5db')])

        self._build_menu()

        self.main_frame = ttk.Frame(root, padding=20)
        self.main_frame.grid(row=0, column=0, sticky='nsew')
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)

        # Instruction card
        instructions = (
            "This program converts Onshape CAM Studio g-code to Shop Sabre WinCNC compatible g-code.\n"
            "To prepare your file for this operation, post from Onshape with these settings:\n"
            "  • Machine = 3-Axis Generic Milling - Fanuc\n"
            "  • Fixed Cycles = All options turned OFF\n"
            "  • Setup -> Position Type = Stock box point"
        )
        info_card = ttk.Frame(self.main_frame, style='Card.TFrame', padding=15)
        info_card.grid(row=0, column=0, sticky='ew')
        info_label = ttk.Label(
            info_card,
            text=instructions,
            style='Card.TLabel',
            justify='left',
            anchor='w',
            wraplength=580
        )
        info_label.grid(row=0, column=0, sticky='w')

        # File selection card
        file_card = ttk.Frame(self.main_frame, style='Card.TFrame', padding=15)
        file_card.grid(row=1, column=0, pady=15, sticky='ew')
        file_card.columnconfigure(1, weight=1)

        ttk.Label(file_card, text='Input G-code file', style='Heading.TLabel').grid(
            row=0, column=0, columnspan=3, sticky='w', pady=(0, 10)
        )
        ttk.Label(file_card, text='Location', style='Card.TLabel').grid(row=1, column=0, sticky='w')
        self.input_entry = ttk.Entry(file_card)
        self.input_entry.grid(row=1, column=1, padx=10, sticky='ew')
        ttk.Button(file_card, text='Browse…', command=self.select_input).grid(row=1, column=2)

        ttk.Label(file_card, text='Output file name', style='Card.TLabel').grid(
            row=2, column=0, sticky='w', pady=(12, 0)
        )
        self.output_entry = ttk.Entry(file_card)
        self.output_entry.grid(row=2, column=1, padx=10, pady=(12, 0), sticky='ew')
        self.output_entry.configure(state='readonly')

        # Options card
        options_card = ttk.Frame(self.main_frame, style='Card.TFrame', padding=15)
        options_card.grid(row=2, column=0, sticky='ew')
        options_card.columnconfigure(0, weight=1)

        ttk.Label(options_card, text='Conversion options', style='Heading.TLabel').grid(
            row=0, column=0, sticky='w', pady=(0, 10)
        )

        self.remove_coolant_var = tk.BooleanVar(value=True)
        self.remove_toolchange_var = tk.BooleanVar(value=True)
        self.zero_plane_var = tk.StringVar(value="Top")

        self.remove_coolant_check = ttk.Checkbutton(
            options_card,
            text='Remove mist commands (M7/M9)',
            variable=self.remove_coolant_var
        )
        self.remove_coolant_check.grid(row=1, column=0, sticky='w')

        self.remove_toolchange_check = ttk.Checkbutton(
            options_card,
            text='Remove tool change commands (M6)',
            variable=self.remove_toolchange_var
        )
        self.remove_toolchange_check.grid(row=2, column=0, sticky='w', pady=(5, 0))

        zero_plane_frame = ttk.Frame(options_card, style='Card.TFrame')
        zero_plane_frame.grid(row=3, column=0, pady=(12, 0), sticky='w')
        ttk.Label(zero_plane_frame, text='Zero plane', style='Card.TLabel').grid(row=0, column=0, sticky='w')
        self.zero_plane_menu = ttk.Combobox(
            zero_plane_frame,
            textvariable=self.zero_plane_var,
            values=['Top', 'Bottom', 'Ignore'],
            state='readonly',
            width=10
        )
        self.zero_plane_menu.grid(row=0, column=1, padx=(10, 0))

        # Actions and status
        action_frame = ttk.Frame(self.main_frame, style='TFrame', padding=(0, 15, 0, 0))
        action_frame.grid(row=3, column=0, sticky='ew')
        action_frame.columnconfigure(0, weight=1)

        self.convert_button = ttk.Button(
            action_frame,
            text='Convert File',
            style='Accent.TButton',
            command=self.convert
        )
        self.convert_button.grid(row=0, column=0, pady=(0, 10))

        self.status_var = tk.StringVar()
        self.status_var.set('Select a file to convert.')
        self.status_label = ttk.Label(action_frame, textvariable=self.status_var, style='Status.TLabel', wraplength=560)
        self.status_label.grid(row=1, column=0, sticky='w')

    def _build_menu(self) -> None:
        """Create the application menu bar with customization entry."""

        menubar = tk.Menu(self.root)
        customize_menu = tk.Menu(menubar, tearoff=0)
        customize_menu.add_command(label='Machine Settings…', command=self.open_customize_dialog)
        customize_menu.add_command(label='Output Settings…', command=self.open_output_settings_dialog)
        menubar.add_cascade(label='Customize', menu=customize_menu)
        self.root.config(menu=menubar)
        self.menubar = menubar

    def select_input(self) -> None:
        """Handle the file selection dialog for the input file."""
        path = filedialog.askopenfilename(
            title='Select Onshape G-code file',
            filetypes=[
                ('G-code Files', '*.nc *.tap *.gcode *.txt'),
                ('All Files', '*.*')
            ]
        )
        if not path:
            return
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, path)
        out_path = self._derive_output_path(path)
        self._set_output_entry(out_path)
        self.status_var.set('Ready to convert.')

    def _set_output_entry(self, value: str) -> None:
        self.output_entry.configure(state='normal')
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, value)
        self.output_entry.configure(state='readonly')

    def _derive_output_path(self, input_path: str) -> str:
        directory = self.settings.output_directory or os.path.dirname(input_path)
        directory = os.path.abspath(os.path.expanduser(directory))
        base_name = os.path.basename(input_path)
        name_root, _ = os.path.splitext(base_name)
        ext = '.tap'
        value = self.settings.output_name_value or ''
        if self.settings.output_name_mode == 'suffix':
            file_name = f"{name_root}{value}{ext}"
        else:
            file_name = f"{value}{name_root}{ext}"
        return os.path.join(directory, file_name)

    def _update_output_entry_for_current_input(self) -> None:
        input_path = self.input_entry.get().strip()
        if not input_path:
            return
        derived = self._derive_output_path(input_path)
        self._set_output_entry(derived)

    def convert(self) -> None:
        """Perform the conversion when the Convert button is clicked."""
        input_path = self.input_entry.get().strip()
        output_path = self.output_entry.get().strip()
        zero_plane = self.zero_plane_var.get().strip().title() or "Top"

        if not input_path:
            messagebox.showerror('Error', 'No input file selected.')
            return
        if not os.path.isfile(input_path):
            messagebox.showerror('Error', 'Input file does not exist.')
            return
        try:
            self.status_var.set('Checking file...')
            self.root.update_idletasks()

            # Read the input file for analysis
            with open(input_path, 'r', encoding='utf-8', errors='ignore') as f_in:
                in_lines = f_in.readlines()

            contains_mist = any(
                re.search(r'\bM7\b', ln, re.IGNORECASE) or re.search(r'\bM9\b', ln, re.IGNORECASE)
                for ln in in_lines
            )
            if contains_mist and not self.remove_coolant_var.get() and self.settings.mist_port is None:
                self.status_var.set('Conversion aborted: mist port not set.')
                messagebox.showerror(
                    'Mister Port Required',
                    'Mist commands were detected but no mister port is configured.\n'
                    'Set your mister port in Machine Settings before converting.'
                )
                return

            # Detect potential SETUP -> POSITION TYPE / Z-zero issues BEFORE conversion.
            if zero_plane != "Ignore" and detect_z_zero_issue(in_lines, zero_plane=zero_plane):
                self.status_var.set('Conversion aborted due to Z-zero / POSITION TYPE error.')
                if zero_plane == "Top":
                    msg = (
                        'Z-zero appears to be incorrect for a TOP-of-stock zero.\n'
                        'No cutting moves go below Z0 after the spindle turns on.\n\n'
                        'Be sure to use:\n'
                        '  Setup -> Position Type = Stock box point\n'
                        'with Z0 at the TOP of the stock in Onshape CAM,\n'
                        'then repost your file and try again.'
                    )
                else:  # Bottom
                    msg = (
                        'Z-zero appears to be incorrect for a BOTTOM-of-part zero.\n'
                        'Cutting moves go below Z0 after the spindle turns on,\n'
                        'which suggests the zero plane or Position Type is wrong.\n\n'
                        'Be sure to use:\n'
                        '  Setup -> Position Type = Stock box point\n'
                        'with Z0 at the BOTTOM of the part in Onshape CAM,\n'
                        'then repost your file and try again.'
                    )
                messagebox.showerror('Z-Zero / POSITION TYPE Error Detected', msg)
                return

            # If no issue detected, perform conversion
            self.status_var.set('Converting...')
            self.root.update_idletasks()
            convert_file(
                input_path,
                output_path,
                remove_coolant=self.remove_coolant_var.get(),
                remove_toolchange=self.remove_toolchange_var.get(),
                mist_port=self.settings.mist_port
            )

            self.status_var.set(f'Conversion complete: {output_path}')
            messagebox.showinfo('Conversion Complete', f'Converted file saved to:\n{output_path}')
        except Exception as e:
            self.status_var.set('Conversion failed.')
            messagebox.showerror('Error', f'An error occurred during conversion:\n{e}')

    def open_customize_dialog(self) -> None:
        """Open the customization dialog for coolant/tool change parameters."""

        if self.settings_window and tk.Toplevel.winfo_exists(self.settings_window):
            self.settings_window.lift()
            self.settings_window.focus_set()
            return

        win = tk.Toplevel(self.root)
        win.title('Customize Machine Settings')
        win.configure(bg='#f4f6fb')
        win.resizable(False, False)
        win.transient(self.root)
        self.settings_window = win
        self.settings_window.protocol('WM_DELETE_WINDOW', self._close_settings_window)

        container = ttk.Frame(win, padding=20)
        container.grid(row=0, column=0, sticky='nsew')

        card = ttk.Frame(container, style='Card.TFrame', padding=15)
        card.grid(row=0, column=0, sticky='ew')
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text='Mister Port', style='Heading.TLabel').grid(row=0, column=0, columnspan=2, sticky='w')

        ttk.Label(card, text='Mist (M7/M9)', style='Card.TLabel').grid(row=1, column=0, sticky='w', pady=(8, 0))
        self.mist_output_var = tk.StringVar(value='' if self.settings.mist_port is None else str(self.settings.mist_port))
        ttk.Entry(card, textvariable=self.mist_output_var).grid(row=1, column=1, padx=(10, 0), pady=(8, 0), sticky='ew')
        ttk.Label(
            card,
            text='WinCNC mister port used with M11C<port> (on) / M12C<port> (off).',
            style='Card.TLabel',
            wraplength=360,
        ).grid(row=2, column=0, columnspan=2, sticky='w')

        ttk.Label(
            card,
            text='Leave the port blank to disable mist conversion. Configure it when your machine wiring is known.',
            style='Card.TLabel',
            wraplength=360,
        ).grid(row=3, column=0, columnspan=2, sticky='w', pady=(12, 0))

        button_frame = ttk.Frame(container, padding=(0, 15, 0, 0))
        button_frame.grid(row=1, column=0, sticky='ew')
        button_frame.columnconfigure(0, weight=1)

        ttk.Button(
            button_frame,
            text='Save Settings',
            style='Accent.TButton',
            command=self._save_settings_from_dialog
        ).grid(row=0, column=0, sticky='ew')

        ttk.Button(
            button_frame,
            text='Cancel',
            command=self._close_settings_window
        ).grid(row=1, column=0, sticky='ew', pady=(10, 0))

    def _close_settings_window(self) -> None:
        if self.settings_window:
            self.settings_window.destroy()
            self.settings_window = None

    def open_output_settings_dialog(self) -> None:
        """Open dialog for configuring output directory and file naming."""

        if self.output_settings_window and tk.Toplevel.winfo_exists(self.output_settings_window):
            self.output_settings_window.lift()
            self.output_settings_window.focus_set()
            return

        win = tk.Toplevel(self.root)
        win.title('Customize Output Settings')
        win.configure(bg='#f4f6fb')
        win.resizable(False, False)
        win.transient(self.root)
        self.output_settings_window = win
        self.output_settings_window.protocol('WM_DELETE_WINDOW', self._close_output_settings_window)

        container = ttk.Frame(win, padding=20)
        container.grid(row=0, column=0, sticky='nsew')

        card = ttk.Frame(container, style='Card.TFrame', padding=15)
        card.grid(row=0, column=0, sticky='ew')
        card.columnconfigure(1, weight=1)

        ttk.Label(card, text='Default Output Location', style='Heading.TLabel').grid(
            row=0, column=0, columnspan=3, sticky='w'
        )
        ttk.Label(card, text='Folder', style='Card.TLabel').grid(row=1, column=0, sticky='w', pady=(8, 0))
        self.output_dir_var = tk.StringVar(value=self.settings.output_directory or '')
        dir_entry = ttk.Entry(card, textvariable=self.output_dir_var)
        dir_entry.grid(row=1, column=1, padx=(10, 0), pady=(8, 0), sticky='ew')
        ttk.Button(card, text='Browse…', command=self._browse_output_directory).grid(
            row=1, column=2, padx=(10, 0), pady=(8, 0)
        )
        ttk.Label(
            card,
            text='Leave blank to save next to the selected input file.',
            style='Card.TLabel',
            wraplength=360,
        ).grid(row=2, column=0, columnspan=3, sticky='w')

        ttk.Label(card, text='File Name Format', style='Heading.TLabel').grid(
            row=3, column=0, columnspan=3, sticky='w', pady=(20, 0)
        )
        ttk.Label(card, text='Mode', style='Card.TLabel').grid(row=4, column=0, sticky='w', pady=(8, 0))
        self.output_mode_var = tk.StringVar(
            value='Prefix' if self.settings.output_name_mode == 'prefix' else 'Suffix'
        )
        ttk.Combobox(
            card,
            textvariable=self.output_mode_var,
            values=['Prefix', 'Suffix'],
            state='readonly',
            width=10,
        ).grid(row=4, column=1, padx=(10, 0), pady=(8, 0), sticky='w')

        ttk.Label(card, text='Text', style='Card.TLabel').grid(row=5, column=0, sticky='w', pady=(8, 0))
        self.output_value_var = tk.StringVar(value=self.settings.output_name_value)
        ttk.Entry(card, textvariable=self.output_value_var).grid(
            row=5, column=1, columnspan=2, padx=(10, 0), pady=(8, 0), sticky='ew'
        )
        ttk.Label(
            card,
            text='Example: prefix of "SS23_" makes SS23_part.tap; suffix of "_WIN" makes part_WIN.tap.',
            style='Card.TLabel',
            wraplength=360,
        ).grid(row=6, column=0, columnspan=3, sticky='w', pady=(4, 0))

        button_frame = ttk.Frame(container, padding=(0, 15, 0, 0))
        button_frame.grid(row=1, column=0, sticky='ew')
        button_frame.columnconfigure(0, weight=1)

        ttk.Button(
            button_frame,
            text='Save Output Settings',
            style='Accent.TButton',
            command=self._save_output_settings
        ).grid(row=0, column=0, sticky='ew')

        ttk.Button(
            button_frame,
            text='Cancel',
            command=self._close_output_settings_window
        ).grid(row=1, column=0, sticky='ew', pady=(10, 0))

    def _browse_output_directory(self) -> None:
        directory = filedialog.askdirectory(title='Select output directory')
        if not directory:
            return
        self.output_dir_var.set(directory)

    def _close_output_settings_window(self) -> None:
        if self.output_settings_window:
            self.output_settings_window.destroy()
            self.output_settings_window = None

    def _save_output_settings(self) -> None:
        directory = self.output_dir_var.get().strip()
        normalized_dir = os.path.abspath(os.path.expanduser(directory)) if directory else None
        if normalized_dir and not os.path.isdir(normalized_dir):
            messagebox.showerror('Invalid Directory', 'The selected directory does not exist. Create it first or leave blank.')
            return

        mode = self.output_mode_var.get().strip().lower()
        mode = 'prefix' if mode not in ('prefix', 'suffix') else mode
        value = self.output_value_var.get()

        self.settings.output_directory = normalized_dir
        self.settings.output_name_mode = mode
        self.settings.output_name_value = value

        try:
            self.settings.save()
        except OSError as exc:
            messagebox.showerror('Save Failed', f'Unable to save settings:\n{exc}')
            return

        self._update_output_entry_for_current_input()
        messagebox.showinfo('Settings Saved', 'Output settings saved successfully.')
        self._close_output_settings_window()

    def _parse_channel_value(self, raw: str, label: str) -> Optional[int]:
        value = (raw or '').strip()
        if not value:
            return None
        try:
            parsed = int(value)
        except ValueError:
            raise ValueError(f'{label} must be a positive integer or left blank.')
        if parsed <= 0:
            raise ValueError(f'{label} must be greater than zero.')
        return parsed

    def _save_settings_from_dialog(self) -> None:
        try:
            mist = self._parse_channel_value(self.mist_output_var.get(), 'Mist port')
        except ValueError as exc:
            messagebox.showerror('Invalid Value', str(exc))
            return

        self.settings.mist_port = mist
        try:
            self.settings.save()
        except OSError as exc:
            messagebox.showerror('Save Failed', f'Unable to save settings:\n{exc}')
            return

        messagebox.showinfo('Settings Saved', 'Machine settings saved successfully.')
        self._close_settings_window()


def main() -> None:
    root = tk.Tk()
    ConverterGUI(root)
    root.resizable(False, False)
    root.mainloop()


if __name__ == '__main__':
    main()
