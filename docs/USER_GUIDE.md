# SystemVerilog FSM Generator User Guide

## 1. Purpose

The tool converts an XLSX transition table into synthesizable SystemVerilog. It is intended
for control FSMs used by any IP block, including ADC, eFuse, reset controllers,
protocol engines, sequencers, power controllers, and timeout handlers.

It generates:

- Module input and output ports.
- State encoding and state registers.
- Next-state combinational logic.
- Combinational output decode.
- Illegal-state recovery to the reset state.
- An optional shared state timer for cycle-based transitions.
- `logic`, typed state `enum`, `always_ff`, `always_comb`, and `unique case`.
- Optional simulation assertions for unknown or illegal states.
- Optional synchronous reset and clock-enable behavior.

It does not generate arbitrary datapaths such as memories, accumulators,
averaging logic, asynchronous synchronizers, or protocol-specific data storage.
Keep those blocks in handwritten RTL and connect their status signals to the
generated FSM.

## 2. Installation

Python 3.8 or newer and `openpyxl` are required for the recommended XLSX workflow:

```sh
python3 -m pip install openpyxl
```

Verify the installation:

```sh
python3 fsm_gen.py --help
```

## 3. Recommended workflow

1. Copy `examples/template.xlsx`, or generate a workbook with
   `make template TEMPLATE_OUT=my_fsm.xlsx`.
2. Fill the `Config`, `Inputs`, and `Outputs` sheets.
3. Add states and transitions in their dedicated sheets.
4. Add one `default` row as the final branch of each state.
5. Add `timer_limit` only to states that need a cycle delay.
6. Generate RTL into a separate output directory.
7. Compile, lint, and simulate the generated module before integration.
8. Keep the XLSX workbook as the source of truth; do not manually edit generated RTL.

## 4. XLSX workbook format

Each workbook separates information into six focused sheets:

| Sheet | What the user edits |
| --- | --- |
| `Guide` | Read-only workflow and entry rules |
| `Config` | Module, clock, reset, reset state, and timer settings |
| `Inputs` | Input port name, width, and description |
| `Outputs` | Output port name, width, default value, and description |
| `States` | State name, optional encoding, timer limit, and description |
| `Transitions` | Priority-ordered conditions, destinations, and output overrides |

Yellow cells are intended for user input. Header rows are frozen, filters are
enabled, and state fields use dropdown lists sourced from the `States` sheet.

### Config sheet

| Key | Required | Meaning | Example |
| --- | --- | --- | --- |
| `module_name` | Recommended | Generated module name | `uart_ctrl_fsm` |
| `clock` | No | Clock port; default `clk` | `pclk` |
| `reset` | No | Reset port; default `rst_n` | `preset_n` |
| `language` | No | `systemverilog` (default) or `verilog` | `systemverilog` |
| `reset_active_low` | No | `true` or `false`; default `true` | `true` |
| `reset_type` | No | `async` (default) or `sync` | `async` |
| `reset_state` | No | State after reset; default is first state | `IDLE` |
| `encoding` | No | Automatic `binary`, `onehot`, or `gray` encoding | `onehot` |
| `timer_width` | No | Shared timer width; default 32 | `8` |
| `clock_enable` | No | Optional enable input; holds state and timer when low | `enable` |
| `assertions` | No | Emit simulation-only SV assertions; default `true` | `true` |

### Inputs and Outputs sheets

Inputs use `name`, `width`, and `description`. Outputs additionally use a
`default` column. Width `1` generates a scalar; width `4` generates `[3:0]`.
Descriptions are for users and are ignored by the RTL generator.

### States sheet

| Column | Required | Meaning | Example |
| --- | --- | --- | --- |
| `state` | Yes | State name | `WAIT_EOC` |
| `code` | No | Explicit state encoding | `4'b0010` |
| `timer_limit` | No | Cycles spent in this state | `16` or `wait_cycles` |
| `description` | No | Human-readable purpose | `Wait for ADC EOC` |

### Transitions sheet

| Column | Required | Meaning | Example |
| --- | --- | --- | --- |
| `state` | Yes | Current state | `WAIT_EOC` |
| `condition` | Yes | Verilog condition or `default` | `eoc && !abort` |
| `next_state` | Yes | Destination state | `CAPTURE` |
| `outputs` | No | Output overrides for this branch | `busy=1'b1;cs=1'b1` |
| `description` | No | Human-readable explanation | `Abort immediately` |

Conditions are emitted in worksheet row order as `if`, `else if`, then `else`. Place the
highest-priority condition first. Use at most one `default` row per state.

Use semicolons between output assignments. Do not merge cells or rename sheets
and header columns.

## 5. Output behavior

All generated outputs are combinational. At the start of the combinational
block, the generator assigns every output its value from `output_defaults`.
The selected transition branch then applies its `outputs` overrides.

Therefore, repeat a state-level output on every branch that needs it:

| state | condition | next_state | outputs |
| --- | --- | --- | --- |
| RUN | abort | IDLE |  |
| RUN | complete | DONE | `busy=1'b1;enable=1'b1` |
| RUN | default | RUN | `busy=1'b1;enable=1'b1` |

In this example, `abort` intentionally uses the defaults and immediately drops
`busy` and `enable`.

## 6. State timing

The generator uses one shared counter for all timed states. A state becomes
timed when any row for that state contains `timer_limit`.

Example with a runtime-programmable delay:

| state | condition | next_state | outputs | timer_limit |
| --- | --- | --- | --- | --- |
| IDLE | start | WAIT | `busy=1'b1` |  |
| IDLE | default | IDLE |  |  |
| WAIT | abort | IDLE |  | delay_cycles |
| WAIT | timer_done | DONE | `busy=1'b1` |  |
| WAIT | default | WAIT | `busy=1'b1` |  |
| DONE | default | IDLE | `done=1'b1` |  |

Timing semantics:

- The entry cycle counts as the first cycle in the state.
- `timer_limit=1` keeps the state for one cycle.
- `timer_limit=4` keeps the state for four cycles.
- Zero is treated like one to avoid subtraction underflow.
- The counter resets whenever the FSM changes state.
- States without `timer_limit` keep the counter inactive and cleared.
- A runtime limit expression must remain stable while its state is active.
- The limit must fit within `timer_width`.

The generated internal names default to:

- `state_timer`
- `timer_active`
- `timer_done`

They may be changed with the optional metadata columns
`timer_counter_name`, `timer_active_name`, and `timer_done_name`.

## 7. Generate from XLSX

Using Make:

```sh
make gen INPUT=examples/timed.xlsx OUT=generated_timed
```

Using Python directly:

```sh
python3 fsm_gen.py examples/timed.xlsx -o generated_timed
python3 fsm_gen.py examples/timed.xlsx --print
```

When `OUT` is a directory, the default filename is `<module_name>.sv`. Pass a
`.sv` path to choose the complete output filename. For a legacy toolchain, use:

```sh
make gen INPUT=examples/timed.xlsx OUT=generated_v LANGUAGE=verilog
```

Compatibility mode writes `<module_name>.v` and uses Verilog-2001 constructs.

## 8. Create a new workbook

Create a workbook:

```sh
make template TEMPLATE_OUT=my_fsm.xlsx
```

This creates `my_fsm.xlsx` with the six user-focused sheets described above,
including dropdowns, filters, editable-cell colors, example rows, and comments.

Generate from it:

```sh
make gen INPUT=my_fsm.xlsx OUT=generated_xlsx
```

The legacy CLI option `--new-adc-template` remains accepted as an alias for
`--new-template`.

## 9. Custom RTL template

For the recommended output, start from `templates/fsm.sv.tpl` and
keep these placeholders:

```text
{{module_name}}
{{port_declarations}}
{{state_declarations}}
{{state_register_block}}
{{timer_declarations}}
{{timer_block}}
{{comb_block}}
{{assertion_block}}
```

Generate with the custom template:

```sh
make gen INPUT=examples/timed.xlsx \
  TEMPLATE=templates/fsm.sv.tpl OUT=generated_custom
```

Timed FSMs require both `{{timer_declarations}}` and `{{timer_block}}`.
The Verilog-2001 compatibility template is `templates/fsm.v.tpl` and uses
the legacy `{{state_localparams}}`, `{{state_reg_name}}`, and
`{{next_state_name}}` placeholders.

## 10. Validation performed by the tool

Generation stops with an error for:

- Missing transitions or states.
- Invalid RTL identifiers.
- Unknown reset or destination states.
- Assignments to undeclared outputs.
- Duplicate transition conditions.
- More than one `default` branch in a state.
- Conflicting `state_code` or `timer_limit` values for one state.
- Invalid `timer_width`.
- Invalid language, reset type, or boolean configuration.
- Duplicate port names.
- Timer internal-name collisions.
- Missing timer placeholders in a custom template.

Generation success only proves that the table is structurally valid. Always
compile and simulate the generated RTL as part of the target IP verification.

## 11. Integration pattern

For complex IP, split control from datapath:

```text
Handwritten datapath/status logic
  ├── edge detection
  ├── counters unrelated to state dwell time
  ├── memories/FIFOs
  ├── arithmetic and accumulation
  └── status conditions
             │
             ▼
Generated FSM control
  ├── state transitions
  ├── state dwell timer
  └── control/status outputs
```

For the ADC example, the generated FSM can control state transitions,
`CS`, `SAMPCLK`, `BUSY`, `READY`, and `DONE`. EOC edge detection, sample RAM,
sum, and average remain handwritten datapath logic.

The complete ADC transition table is kept as an advanced reference at
`examples/adc_timed.xlsx`.

## 12. Troubleshooting

### `timer_done` is undeclared

Add `timer_limit` to that state. If `timer_done` is meant to be an external
signal instead, declare it in `inputs` and do not use the same name for the
internal timer.

### State exits one cycle too early or late

Remember that the entry cycle is cycle one. Check that the programmed limit is
stable and fits within `timer_width`.

### Output drops during a transition

Output overrides apply per branch. Repeat state outputs on every branch that
must keep them asserted.

### XLSX cannot be read

Install `openpyxl` and keep the standard multi-sheet workbook names:

```sh
python3 -m pip install openpyxl
make gen INPUT=my_fsm.xlsx OUT=generated
```

### Clean generated artifacts

```sh
make clean
```

The clean target removes generated RTL, Python caches, and simulation outputs;
it does not remove source XLSX specifications.
