# Verilog FSM Generator

Generate a Verilog FSM from one CSV table or one sheet in an XLSX workbook.

## Quick Start

```sh
cd /FSM_Generator
make gen
```

Default output:

```text
generated/adc_controller_fsm.v
```

Clean generated files:

```sh
make clean
```

## One-Table CSV

For a new FSM, copy or edit the blank template:

```text
examples/adc_controller.csv
```

The table format is:

| module_name | clock | reset | reset_active_low | reset_state | inputs | output_defaults | state | condition | next_state | outputs |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| my_fsm | clk | rst_n | true | IDLE | start;done | out_valid=1'b0;busy=1'b0 | IDLE | start | RUN | busy=1'b1 |
|  |  |  |  |  |  |  | IDLE | default | IDLE |  |
|  |  |  |  |  |  |  | RUN | done | DONE | out_valid=1'b1 |
|  |  |  |  |  |  |  | RUN | default | RUN | busy=1'b1 |
|  |  |  |  |  |  |  | DONE | default | IDLE |  |

In the actual CSV file, each row is saved as comma-separated values. Blank cells are allowed after the first data row.

Only the first data row needs `module_name`, `clock`, `reset`, `inputs`, and `output_defaults`. The rows after that can leave those columns blank.

Generate from your filled CSV:

```sh
make gen INPUT=examples/adc_controller.csv
```

## Examples

Generate the ADC controller example:

```sh
make gen INPUT=examples/adc_controller_example.csv
```

Generate the traffic-light example:

```sh
make gen SHEET=traffic_light_example OUT=generated_traffic
```

`SHEET=name` is a shortcut for `examples/name.csv` when that CSV exists.

## XLSX Template

CSV files do not have sheets. If you want a workbook with a guide sheet plus a data sheet, create the XLSX template:

```sh
make template
```

This creates:

```text
examples/adc_controller_template.xlsx
```

It contains:

- Sheet `template`: guide and column meanings.
- Sheet `adc_controller`: example FSM data used by the generator.

Generate from the workbook sheet:

```sh
make gen INPUT=examples/adc_controller_template.xlsx SHEET=adc_controller OUT=generated_xlsx
```

## Columns

Required transition columns:

- `state`: current state
- `condition`: transition condition, or `default`
- `next_state`: next state
- `outputs`: output assignments for that branch

Optional config columns:

- `module_name`: generated Verilog module name
- `clock`: clock port name, default `clk`
- `reset`: reset port name, default `rst_n`
- `reset_active_low`: `true` or `false`, default `true`
- `reset_state`: reset state. If blank, first state is used.
- `inputs`: semicolon-separated input list, for example `start;timer_done;mode:2`
- `output_defaults`: semicolon-separated output defaults, for example `done=1'b0;data:8=8'd0`
- `state_code`: optional explicit encoding for the current state, for example `3'b010`

Output assignments are separated by `;`:

```text
green=1'b1;timer_load=1'b1
```

## Coverage Behavior

The generator covers the transition table by:

- Inferring every state from `state` and `next_state`.
- Emitting one `case` item for every state.
- Emitting each transition row for that state in CSV order.
- Emitting an `else` branch for every state. If no `default` row exists, it holds the current state.
- Emitting a top-level `default` case that sends illegal states to `reset_state`.
- Failing generation if a transition references an unknown output.

## Commands

```sh
python3 fsm_gen.py examples/adc_controller_example.csv -o generated
python3 fsm_gen.py examples/adc_controller_example.csv --print
python3 fsm_gen.py examples/adc_controller_template.xlsx --sheet adc_controller -o generated_xlsx
python3 fsm_gen.py --new-adc-template examples/adc_controller_template.xlsx
```

Use a custom Verilog template:

```sh
python3 fsm_gen.py examples/traffic_light_example.csv --template templates/verilog_fsm.tpl -o generated
```

## Excel Dependency

XLSX support needs `openpyxl`:

```sh
python3 -m pip install openpyxl
```
