# SystemVerilog FSM Generator

Generate synthesizable SystemVerilog FSMs from Excel (`.xlsx`) workbooks.
SystemVerilog (`.sv`) is the default output; Verilog-2001 is also supported.

## Quick start

Requirements: Python 3.8+ and `openpyxl`.

```sh
python3 -m pip install openpyxl
make gen
```

Default output:

```text
generated/timed_controller_fsm.sv
```

## Create your own FSM

1. Create a workbook template:

```sh
make template TEMPLATE_OUT=my_fsm.xlsx
```

2. Open `my_fsm.xlsx` and edit the `Config`, `Inputs`, `Outputs`, `States`, and
   `Transitions` sheets.

3. Generate the RTL:

```sh
make gen INPUT=my_fsm.xlsx OUT=generated
```

## State timing

Set `timer_limit` in the `States` sheet, then use `timer_done` as a transition
condition:

```text
WAIT | timer_done | DONE
WAIT | default    | WAIT
```

`timer_limit` may be a fixed value such as `16` or an input such as
`delay_cycles`.

## Common commands

```sh
# Generate the default example
make gen

# Generate another workbook
make gen INPUT=examples/adc_timed.xlsx OUT=generated_adc

# Generate Verilog-2001 instead of SystemVerilog
make gen INPUT=my_fsm.xlsx OUT=generated_v LANGUAGE=verilog

# Remove generated files and caches
make clean

# Show command help
make help
python3 fsm_gen.py --help
```

Example workbooks are available in `examples/`. For the complete workbook
format, reset options, state encoding, clock enable, timing, and custom
templates, see [docs/USER_GUIDE.md](docs/USER_GUIDE.md).
