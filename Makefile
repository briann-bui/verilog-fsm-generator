PYTHON ?= python3
INPUT ?= examples/timed.xlsx
SHEET ?=
OUT ?= generated
TEMPLATE ?=
LANGUAGE ?=
TEMPLATE_OUT ?= fsm_template.xlsx

SHEET_ARG = $(if $(SHEET),--sheet $(SHEET),)
TEMPLATE_ARG = $(if $(TEMPLATE),--template $(TEMPLATE),)
LANGUAGE_ARG = $(if $(LANGUAGE),--language $(LANGUAGE),)

.PHONY: gen template clean help

gen:
	$(PYTHON) fsm_gen.py $(INPUT) $(SHEET_ARG) $(TEMPLATE_ARG) $(LANGUAGE_ARG) -o $(OUT)

template:
	$(PYTHON) fsm_gen.py --new-template $(TEMPLATE_OUT)

clean:
	rm -rf generated generated_* sim_check sim_check_*.log csrc __pycache__ *.pyc
	find . -type d -name '__pycache__' -prune -exec rm -rf {} +
	rm -rf simv simv.daidir simv.vdb

help:
	@echo "SystemVerilog FSM Generator"
	@echo ""
	@echo "  make gen"
	@echo "      Generate the default timed-controller example into generated/."
	@echo ""
	@echo "  make gen INPUT=<file.xlsx> [SHEET=<sheet>] [OUT=<path>] [LANGUAGE=systemverilog|verilog]"
	@echo "      Generate SystemVerilog by default; legacy Verilog remains supported."
	@echo ""
	@echo "  make template [TEMPLATE_OUT=<file.xlsx>]"
	@echo "      Create a generic XLSX workbook template."
	@echo ""
	@echo "  make clean"
	@echo "      Remove generated RTL, caches, and simulation outputs."
