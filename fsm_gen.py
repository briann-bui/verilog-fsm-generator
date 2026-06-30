#!/usr/bin/env python3
import argparse
import csv
import math
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple


IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_$]*$")


DEFAULT_TEMPLATE = """\
module {{module_name}} (
{{port_declarations}}
);

{{state_localparams}}

reg [STATE_W-1:0] {{state_reg_name}};
reg [STATE_W-1:0] {{next_state_name}};

{{state_register_block}}

{{comb_block}}

endmodule
"""


@dataclass
class Field:
    name: str
    width: str = "1"
    default: str = "0"

    @property
    def decl_width(self) -> str:
        width = str(self.width).strip()
        if width in ("", "1", "1'b1"):
            return ""
        if width.startswith("[") and width.endswith("]"):
            return f" {width}"
        if width.isdigit():
            value = int(width)
            if value <= 1:
                return ""
            return f" [{value - 1}:0]"
        return f" {width}"


@dataclass
class State:
    raw_name: str
    ident: str
    code: Optional[str] = None


@dataclass
class Transition:
    state: str
    condition: str
    next_state: str
    outputs: Dict[str, str] = field(default_factory=dict)
    row_num: int = 0

    @property
    def is_default(self) -> bool:
        return self.condition.strip().lower() in ("", "default", "else", "1", "1'b1", "true")


@dataclass
class Spec:
    meta: Dict[str, str]
    states: List[State]
    inputs: List[Field]
    outputs: List[Field]
    transitions: List[Transition]
    template: str = DEFAULT_TEMPLATE


def die(message: str) -> None:
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def warn(message: str) -> None:
    print(f"WARNING: {message}", file=sys.stderr)


ADC_TEMPLATE_HEADERS = [
    "module_name",
    "clock",
    "reset",
    "reset_active_low",
    "reset_state",
    "inputs",
    "output_defaults",
    "state",
    "condition",
    "next_state",
    "outputs",
]


ADC_TEMPLATE_ROWS = [
    [
        "adc_controller_fsm",
        "clk",
        "rst_n",
        "true",
        "IDLE",
        "start;eoc;timeout;fifo_full",
        "adc_cs_n=1'b1;adc_start=1'b0;sample_valid=1'b0;fifo_wr=1'b0;error=1'b0;busy=1'b0",
        "IDLE",
        "start",
        "START_CONV",
        "adc_cs_n=1'b0;adc_start=1'b1;busy=1'b1",
    ],
    ["", "", "", "", "", "", "", "IDLE", "default", "IDLE", ""],
    ["", "", "", "", "", "", "", "START_CONV", "default", "WAIT_EOC", "adc_cs_n=1'b0;busy=1'b1"],
    ["", "", "", "", "", "", "", "WAIT_EOC", "eoc && !fifo_full", "READ_SAMPLE", "adc_cs_n=1'b0;busy=1'b1"],
    ["", "", "", "", "", "", "", "WAIT_EOC", "eoc && fifo_full", "ERROR", "error=1'b1"],
    ["", "", "", "", "", "", "", "WAIT_EOC", "timeout", "ERROR", "error=1'b1"],
    ["", "", "", "", "", "", "", "WAIT_EOC", "default", "WAIT_EOC", "adc_cs_n=1'b0;busy=1'b1"],
    ["", "", "", "", "", "", "", "READ_SAMPLE", "default", "WRITE_FIFO", "sample_valid=1'b1;busy=1'b1"],
    ["", "", "", "", "", "", "", "WRITE_FIFO", "default", "DONE", "fifo_wr=1'b1;busy=1'b1"],
    ["", "", "", "", "", "", "", "DONE", "default", "IDLE", ""],
    ["", "", "", "", "", "", "", "ERROR", "default", "IDLE", "error=1'b1"],
]


def read_csv_rows(path: Path) -> List[Dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as fp:
        reader = csv.DictReader(fp)
        if not reader.fieldnames:
            return []
        rows = []
        for row in reader:
            normalized = {}
            for key, value in row.items():
                if key is None:
                    continue
                normalized[key.strip().lower()] = "" if value is None else str(value).strip()
            if any(value != "" for value in normalized.values()):
                rows.append(normalized)
        return rows


def read_template_csv(path: Path) -> str:
    lines = []
    with path.open(newline="", encoding="utf-8-sig") as fp:
        reader = csv.reader(fp)
        for row in reader:
            if not row:
                lines.append("")
            else:
                lines.append(row[0])
    return "\n".join(lines)


def read_xlsx_sheets(path: Path) -> Tuple[Dict[str, List[Dict[str, str]]], Optional[str]]:
    try:
        import openpyxl
    except ImportError:
        die("Reading .xlsx needs openpyxl. Install with: python3 -m pip install openpyxl")

    wb = openpyxl.load_workbook(path, data_only=True)
    sheets: Dict[str, List[Dict[str, str]]] = {}
    template = None

    for ws in wb.worksheets:
        sheet_name = ws.title.strip().lower()
        values = list(ws.iter_rows(values_only=True))
        if sheet_name == "template":
            lines = []
            for row in values:
                first = "" if not row or row[0] is None else str(row[0])
                lines.append(first)
            candidate = "\n".join(lines).rstrip() + "\n"
            if "{{port_declarations}}" in candidate or "{{comb_block}}" in candidate:
                template = candidate
            continue

        non_empty = [row for row in values if row and any(cell is not None and str(cell).strip() for cell in row)]
        if not non_empty:
            sheets[sheet_name] = []
            continue

        headers = ["" if cell is None else str(cell).strip().lower() for cell in non_empty[0]]
        rows = []
        for raw_row in non_empty[1:]:
            row = {}
            for idx, header in enumerate(headers):
                if not header:
                    continue
                cell = raw_row[idx] if idx < len(raw_row) else ""
                row[header] = "" if cell is None else str(cell).strip()
            if any(value != "" for value in row.values()):
                rows.append(row)
        sheets[sheet_name] = rows

    return sheets, template


def create_adc_template(path: Path) -> None:
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        die("Creating .xlsx templates needs openpyxl. Install with: python3 -m pip install openpyxl")

    wb = openpyxl.Workbook()
    ws_help = wb.active
    ws_help.title = "template"
    help_rows = [
        ["ADC Controller FSM Template"],
        ["Sheet nay la huong dan. Generator se bo qua sheet nay va doc sheet adc_controller."],
        ["Chi can sua bang trong sheet adc_controller de tao FSM Verilog."],
        [""],
        ["Cot", "Y nghia", "Vi du"],
        ["module_name", "Ten module Verilog. Chi can dien o dong dau.", "adc_controller_fsm"],
        ["clock", "Ten clock input.", "clk"],
        ["reset", "Ten reset input.", "rst_n"],
        ["reset_active_low", "true neu reset active-low, false neu active-high.", "true"],
        ["reset_state", "State reset. Neu de trong se dung state dau tien.", "IDLE"],
        ["inputs", "Danh sach input cach nhau bang dau ;. Bus dung name:width.", "start;eoc;timeout;fifo_full;channel_sel:3"],
        ["output_defaults", "Output default moi chu ky, cach nhau bang dau ;.", "adc_cs_n=1'b1;busy=1'b0"],
        ["state", "Current state.", "WAIT_EOC"],
        ["condition", "Dieu kien chuyen state. Dung default cho nhanh else.", "eoc && !fifo_full"],
        ["next_state", "Next state.", "READ_SAMPLE"],
        ["outputs", "Output override cho branch do, cach nhau bang dau ;.", "sample_valid=1'b1;fifo_wr=1'b1"],
        [""],
        ["Lenh gen"],
        [f"python3 fsm_gen.py {path.name} -o generated"],
    ]
    for row in help_rows:
        ws_help.append(row)

    ws_help["A1"].font = Font(bold=True, size=14)
    ws_help["A1"].fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws_help[5]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E2F0D9")
    for idx, width in enumerate([24, 76, 52], start=1):
        ws_help.column_dimensions[get_column_letter(idx)].width = width

    ws_data = wb.create_sheet("adc_controller")
    ws_data.append(ADC_TEMPLATE_HEADERS)
    for row in ADC_TEMPLATE_ROWS:
        ws_data.append(row)
    for cell in ws_data[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FFF2CC")
    widths = [24, 10, 10, 18, 16, 38, 96, 18, 28, 18, 80]
    for idx, width in enumerate(widths, start=1):
        ws_data.column_dimensions[get_column_letter(idx)].width = width

    for ws in (ws_help, ws_data):
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def load_sheets(path: Path, template_path: Optional[Path], sheet_name: Optional[str]) -> Tuple[Dict[str, List[Dict[str, str]]], str]:
    if path.is_dir():
        sheets: Dict[str, List[Dict[str, str]]] = {}
        template = None
        for csv_path in sorted(path.glob("*.csv")):
            key = csv_path.stem.strip().lower()
            if key == "template":
                template = read_template_csv(csv_path)
            else:
                sheets[key] = read_csv_rows(csv_path)
        if sheet_name:
            key = sheet_name.strip().lower()
            if key not in sheets:
                die(f"Folder input has no CSV sheet named '{sheet_name}'")
            sheets = {"table": sheets[key]}
        if template_path:
            template = template_path.read_text(encoding="utf-8")
        return sheets, template or DEFAULT_TEMPLATE

    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        sheets, template = read_xlsx_sheets(path)
        if sheet_name:
            key = sheet_name.strip().lower()
            if key not in sheets:
                die(f"Workbook has no data sheet named '{sheet_name}'")
            sheets = {"table": sheets[key]}
            if template_path:
                template = template_path.read_text(encoding="utf-8")
            return sheets, template or DEFAULT_TEMPLATE
        known = {"meta", "states", "inputs", "outputs", "transitions", "template"}
        data_sheets = [name for name, rows in sheets.items() if name not in known and rows]
        if "transitions" not in sheets and len(data_sheets) == 1:
            sheets = {"table": sheets[data_sheets[0]]}
        elif "transitions" not in sheets and "table" not in sheets and len(data_sheets) > 1:
            die(f"Workbook has multiple data sheets {data_sheets}. Keep one data sheet or use a sheet named transitions.")
        if template_path:
            template = template_path.read_text(encoding="utf-8")
        return sheets, template or DEFAULT_TEMPLATE

    if suffix == ".csv":
        rows = read_csv_rows(path)
        if rows and "sheet" in rows[0]:
            sheets: Dict[str, List[Dict[str, str]]] = {}
            for row in rows:
                sheet = row.pop("sheet", "").strip().lower()
                if sheet:
                    sheets.setdefault(sheet, []).append(row)
            if sheet_name:
                key = sheet_name.strip().lower()
                if key not in sheets:
                    die(f"CSV input has no sheet column value named '{sheet_name}'")
                sheets = {"table": sheets[key]}
        else:
            sheets = {"table": rows}
        template = template_path.read_text(encoding="utf-8") if template_path else DEFAULT_TEMPLATE
        return sheets, template

    die(f"Unsupported input type: {path}")


def pick(row: Dict[str, str], *keys: str, default: str = "") -> str:
    for key in keys:
        value = row.get(key.lower(), "")
        if value != "":
            return value
    return default


def sanitize_ident(name: str, prefix: str = "S") -> str:
    candidate = re.sub(r"[^A-Za-z0-9_$]", "_", name.strip())
    if not candidate:
        candidate = prefix
    if not re.match(r"^[A-Za-z_]", candidate):
        candidate = f"{prefix}_{candidate}"
    return candidate


def require_ident(name: str, context: str) -> str:
    if not IDENT_RE.match(name):
        die(f"Invalid Verilog identifier '{name}' in {context}")
    return name


def parse_meta(rows: List[Dict[str, str]], fallback_module: str) -> Dict[str, str]:
    meta = {
        "module_name": fallback_module,
        "clock": "clk",
        "reset": "rst_n",
        "reset_active_low": "true",
        "reset_state": "",
        "state_reg_name": "state",
        "next_state_name": "next_state",
        "encoding": "binary",
        "output_file": "",
    }
    for row in rows:
        key = pick(row, "key", "name")
        value = pick(row, "value")
        if key:
            meta[key.strip().lower()] = value
    return meta


def parse_fields(rows: List[Dict[str, str]], kind: str) -> List[Field]:
    fields = []
    for idx, row in enumerate(rows, start=2):
        name = pick(row, "name", kind)
        if not name:
            continue
        require_ident(name, f"{kind} row {idx}")
        fields.append(
            Field(
                name=name,
                width=pick(row, "width", "bits", default="1"),
                default=pick(row, "default", "reset", default="0"),
            )
        )
    return fields


def split_items(text: str) -> List[str]:
    return [item.strip() for item in re.split(r"[;\n]", text) if item.strip()]


def parse_input_list(text: str) -> List[Field]:
    fields = []
    for item in split_items(text):
        if ":" in item:
            name, width = item.split(":", 1)
        else:
            name, width = item, "1"
        name = name.strip()
        width = width.strip() or "1"
        require_ident(name, "single-table inputs")
        fields.append(Field(name=name, width=width, default="0"))
    return fields


def parse_output_list(text: str) -> List[Field]:
    fields = []
    for item in split_items(text):
        left, default = (item.split("=", 1) + ["0"])[:2] if "=" in item else (item, "0")
        if ":" in left:
            name, width = left.split(":", 1)
        else:
            name, width = left, "1"
        name = name.strip()
        width = width.strip() or "1"
        default = default.strip() or "0"
        require_ident(name, "single-table output_defaults")
        fields.append(Field(name=name, width=width, default=default))
    return fields


def first_non_empty(rows: List[Dict[str, str]], *keys: str) -> str:
    for row in rows:
        value = pick(row, *keys)
        if value:
            return value
    return ""


def build_single_table_spec(rows: List[Dict[str, str]], fallback_module: str, template: str) -> Spec:
    meta = parse_meta([], fallback_module)
    for key in (
        "module_name",
        "clock",
        "reset",
        "reset_active_low",
        "reset_state",
        "state_reg_name",
        "next_state_name",
        "output_file",
    ):
        value = first_non_empty(rows, key)
        if value:
            meta[key] = value

    transitions = parse_transitions(rows)

    state_rows = []
    seen_state_rows = set()
    for row in rows:
        state = pick(row, "state", "current_state", "current")
        code = pick(row, "state_code", "code", "encoding")
        if state and state not in seen_state_rows:
            seen_state_rows.add(state)
            state_rows.append({"state": state, "code": code})
    states = parse_states(state_rows, transitions)

    input_text = first_non_empty(rows, "inputs", "input_list")
    output_text = first_non_empty(rows, "output_defaults", "outputs_default", "output_list")
    inputs = parse_input_list(input_text)
    outputs = parse_output_list(output_text)

    if not outputs:
        inferred = []
        seen_outputs = set()
        for tr in transitions:
            for name in tr.outputs:
                if name not in seen_outputs:
                    seen_outputs.add(name)
                    inferred.append(Field(name=name, width="1", default="0"))
        outputs = inferred

    return Spec(meta=meta, states=states, inputs=inputs, outputs=outputs, transitions=transitions, template=template)


def parse_output_assignments(text: str, row_num: int) -> Dict[str, str]:
    result = {}
    if not text.strip():
        return result
    parts = split_items(text)
    for part in parts:
        item = part.strip()
        if not item:
            continue
        if "=" not in item:
            die(f"Transition row {row_num}: output assignment '{item}' must use name=value")
        name, value = item.split("=", 1)
        name = name.strip()
        value = value.strip()
        require_ident(name, f"transition row {row_num} output assignment")
        if not value:
            die(f"Transition row {row_num}: output '{name}' has empty value")
        result[name] = value
    return result


def parse_transitions(rows: List[Dict[str, str]]) -> List[Transition]:
    transitions = []
    for idx, row in enumerate(rows, start=2):
        state = pick(row, "state", "current_state", "current")
        next_state = pick(row, "next_state", "next")
        condition = pick(row, "condition", "cond", "when", default="default")
        outputs = pick(row, "outputs", "assign", "assignments", default="")
        if not state and not next_state:
            continue
        if not state or not next_state:
            die(f"Transition row {idx}: both state and next_state are required")
        transitions.append(
            Transition(
                state=state,
                condition=condition,
                next_state=next_state,
                outputs=parse_output_assignments(outputs, idx),
                row_num=idx,
            )
        )
    return transitions


def parse_states(rows: List[Dict[str, str]], transitions: List[Transition]) -> List[State]:
    raw_states = []
    codes: Dict[str, str] = {}

    for idx, row in enumerate(rows, start=2):
        name = pick(row, "state", "name")
        if not name:
            continue
        raw_states.append(name)
        code = pick(row, "code", "encoding")
        if code:
            codes[name] = code

    if not raw_states:
        for tr in transitions:
            if tr.state not in raw_states:
                raw_states.append(tr.state)
            if tr.next_state not in raw_states:
                raw_states.append(tr.next_state)

    states = []
    used_idents = set()
    for raw in raw_states:
        ident = sanitize_ident(raw)
        if not IDENT_RE.match(ident):
            die(f"Cannot sanitize state name '{raw}' into a valid Verilog identifier")
        if ident in used_idents:
            die(f"Duplicate state identifier after sanitize: {ident}")
        used_idents.add(ident)
        states.append(State(raw_name=raw, ident=ident, code=codes.get(raw)))

    return states


def build_spec(path: Path, template_path: Optional[Path], sheet_name: Optional[str] = None) -> Spec:
    sheets, template = load_sheets(path, template_path, sheet_name)
    fallback_module = sanitize_ident(path.stem if path.is_file() else path.name, "fsm")
    if "table" in sheets:
        return build_single_table_spec(sheets["table"], fallback_module, template)
    transitions = parse_transitions(sheets.get("transitions", []))
    states = parse_states(sheets.get("states", []), transitions)
    meta = parse_meta(sheets.get("meta", []), fallback_module)
    inputs = parse_fields(sheets.get("inputs", []), "input")
    outputs = parse_fields(sheets.get("outputs", []), "output")
    return Spec(meta=meta, states=states, inputs=inputs, outputs=outputs, transitions=transitions, template=template)


def state_lookup(spec: Spec) -> Dict[str, State]:
    lookup = {}
    for state in spec.states:
        lookup[state.raw_name] = state
        lookup[state.ident] = state
    return lookup


def state_width(spec: Spec) -> int:
    coded_width = 0
    for state in spec.states:
        if state.code:
            match = re.match(r"(\d+)'[bhdBHD]", state.code.strip())
            if match:
                coded_width = max(coded_width, int(match.group(1)))
    if coded_width:
        return coded_width
    return max(1, math.ceil(math.log2(max(1, len(spec.states)))))


def assign_state_codes(spec: Spec) -> None:
    width = state_width(spec)
    for idx, state in enumerate(spec.states):
        if not state.code:
            state.code = f"{width}'d{idx}"


def validate_spec(spec: Spec) -> None:
    if not spec.transitions:
        die("No transitions found. Add transitions.csv or a transitions sheet.")
    if not spec.states:
        die("No states found. Add states.csv or state rows in transitions.csv.")

    require_ident(spec.meta["module_name"], "meta.module_name")
    require_ident(spec.meta["clock"], "meta.clock")
    require_ident(spec.meta["reset"], "meta.reset")
    require_ident(spec.meta["state_reg_name"], "meta.state_reg_name")
    require_ident(spec.meta["next_state_name"], "meta.next_state_name")

    lookup = state_lookup(spec)
    reset_state = spec.meta.get("reset_state", "")
    if not reset_state:
        spec.meta["reset_state"] = spec.states[0].raw_name
    elif reset_state not in lookup:
        die(f"reset_state '{reset_state}' is not listed in states")

    output_names = {item.name for item in spec.outputs}
    for tr in spec.transitions:
        if tr.state not in lookup:
            die(f"Transition row {tr.row_num}: state '{tr.state}' is not listed in states")
        if tr.next_state not in lookup:
            die(f"Transition row {tr.row_num}: next_state '{tr.next_state}' is not listed in states")
        for name in tr.outputs:
            if name not in output_names:
                die(f"Transition row {tr.row_num}: output '{name}' is not listed in outputs")

    seen_default = set()
    seen_condition = set()
    for tr in spec.transitions:
        key = lookup[tr.state].ident
        if tr.is_default:
            if key in seen_default:
                die(f"State '{tr.state}' has more than one default transition")
            seen_default.add(key)
        else:
            cond_key = (key, tr.condition)
            if cond_key in seen_condition:
                die(f"State '{tr.state}' has duplicate condition '{tr.condition}'")
            seen_condition.add(cond_key)

    assign_state_codes(spec)


def indent(text: str, spaces: int) -> str:
    pad = " " * spaces
    return "\n".join(pad + line if line else line for line in text.splitlines())


def render_ports(spec: Spec) -> str:
    ports = []
    ports.append(f"  input wire {spec.meta['clock']}")
    ports.append(f"  input wire {spec.meta['reset']}")
    for item in spec.inputs:
        ports.append(f"  input wire{item.decl_width} {item.name}")
    for item in spec.outputs:
        ports.append(f"  output reg{item.decl_width} {item.name}")
    return ",\n".join(ports)


def render_state_localparams(spec: Spec) -> str:
    width = state_width(spec)
    lines = [f"localparam integer STATE_W = {width};"]
    for state in spec.states:
        lines.append(f"localparam [STATE_W-1:0] {state.ident} = {state.code};")
    return "\n".join(lines)


def active_reset_condition(spec: Spec) -> Tuple[str, str]:
    rst = spec.meta["reset"]
    active_low = spec.meta.get("reset_active_low", "true").strip().lower() in ("1", "true", "yes", "low", "active_low")
    if active_low:
        return f"negedge {rst}", f"!{rst}"
    return f"posedge {rst}", rst


def render_state_register_block(spec: Spec, lookup: Dict[str, State]) -> str:
    clk = spec.meta["clock"]
    state_reg = spec.meta["state_reg_name"]
    next_state = spec.meta["next_state_name"]
    reset_state = lookup[spec.meta["reset_state"]].ident
    reset_edge, reset_cond = active_reset_condition(spec)
    return "\n".join(
        [
            f"always @(posedge {clk} or {reset_edge}) begin",
            f"  if ({reset_cond}) begin",
            f"    {state_reg} <= {reset_state};",
            "  end else begin",
            f"    {state_reg} <= {next_state};",
            "  end",
            "end",
        ]
    )


def output_default_lines(spec: Spec, spaces: int = 2) -> List[str]:
    pad = " " * spaces
    return [f"{pad}{item.name} = {item.default};" for item in spec.outputs]


def assignment_lines(spec: Spec, tr: Transition, lookup: Dict[str, State], spaces: int) -> List[str]:
    pad = " " * spaces
    lines = [f"{pad}{spec.meta['next_state_name']} = {lookup[tr.next_state].ident};"]
    for name, value in tr.outputs.items():
        lines.append(f"{pad}{name} = {value};")
    return lines


def render_state_branch(spec: Spec, state: State, transitions: List[Transition], lookup: Dict[str, State]) -> str:
    state_transitions = [tr for tr in transitions if lookup[tr.state].ident == state.ident]
    conditional = [tr for tr in state_transitions if not tr.is_default]
    default_tr = next((tr for tr in state_transitions if tr.is_default), None)
    next_name = spec.meta["next_state_name"]

    lines = [f"    {state.ident}: begin"]
    if conditional:
        for idx, tr in enumerate(conditional):
            keyword = "if" if idx == 0 else "else if"
            lines.append(f"      {keyword} ({tr.condition}) begin")
            lines.extend(assignment_lines(spec, tr, lookup, 8))
            lines.append("      end")
        lines.append("      else begin")
        if default_tr:
            lines.extend(assignment_lines(spec, default_tr, lookup, 8))
        else:
            lines.append(f"        {next_name} = {state.ident};")
        lines.append("      end")
    elif default_tr:
        lines.extend(assignment_lines(spec, default_tr, lookup, 6))
    else:
        lines.append(f"      {next_name} = {state.ident};")
    lines.append("    end")
    return "\n".join(lines)


def render_comb_block(spec: Spec, lookup: Dict[str, State]) -> str:
    state_reg = spec.meta["state_reg_name"]
    next_state = spec.meta["next_state_name"]
    reset_state = lookup[spec.meta["reset_state"]].ident
    lines = ["always @(*) begin"]
    lines.append(f"  {next_state} = {state_reg};")
    lines.extend(output_default_lines(spec, 2))
    lines.append("")
    lines.append(f"  case ({state_reg})")
    for state in spec.states:
        lines.append(render_state_branch(spec, state, spec.transitions, lookup))
    lines.append("    default: begin")
    lines.append(f"      {next_state} = {reset_state};")
    lines.append("    end")
    lines.append("  endcase")
    lines.append("end")
    return "\n".join(lines)


def render(spec: Spec) -> str:
    validate_spec(spec)
    lookup = state_lookup(spec)
    replacements = {
        "module_name": spec.meta["module_name"],
        "port_declarations": render_ports(spec),
        "state_localparams": render_state_localparams(spec),
        "state_reg_name": spec.meta["state_reg_name"],
        "next_state_name": spec.meta["next_state_name"],
        "state_register_block": render_state_register_block(spec, lookup),
        "comb_block": render_comb_block(spec, lookup),
    }
    text = spec.template
    for key, value in replacements.items():
        text = text.replace("{{" + key + "}}", value)
    return text.rstrip() + "\n"


def write_output(spec: Spec, input_path: Path, out_arg: Optional[Path], text: str) -> Path:
    if out_arg:
        if out_arg.is_dir() or str(out_arg).endswith("/") or out_arg.suffix == "":
            out_path = out_arg / f"{spec.meta['module_name']}.v"
        else:
            out_path = out_arg
    elif spec.meta.get("output_file"):
        out_path = input_path.parent / spec.meta["output_file"]
    else:
        out_path = Path.cwd() / f"{spec.meta['module_name']}.v"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(text, encoding="utf-8")
    return out_path


def main(argv: Optional[Iterable[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Generate a Verilog FSM from one CSV table or one XLSX data sheet.")
    parser.add_argument("input", nargs="?", type=Path, help="Spec folder, .csv file, or .xlsx workbook")
    parser.add_argument("-o", "--out", type=Path, help="Output .v file or output directory")
    parser.add_argument("--sheet", help="Data sheet name for .xlsx, folder CSV, or CSV files with a sheet column")
    parser.add_argument("--template", type=Path, help="Optional template file overriding template sheet")
    parser.add_argument("--print", action="store_true", help="Print generated Verilog to stdout")
    parser.add_argument("--new-adc-template", type=Path, help="Create an ADC controller workbook template and exit")
    args = parser.parse_args(argv)

    if args.new_adc_template:
        create_adc_template(args.new_adc_template)
        print(f"Created {args.new_adc_template}")
        return 0

    if not args.input:
        parser.error("input is required unless --new-adc-template is used")

    spec = build_spec(args.input, args.template, args.sheet)
    text = render(spec)
    if args.print:
        print(text, end="")
        return 0
    out_path = write_output(spec, args.input, args.out, text)
    print(f"Generated {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
