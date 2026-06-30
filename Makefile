PYTHON ?= python3
INPUT ?= examples/adc_controller_example.csv
SHEET ?=
OUT ?= generated
TEMPLATE ?=

CSV_FROM_SHEET_EXT = $(if $(filter %.csv,$(SHEET)),examples/$(SHEET),)
CSV_FROM_SHEET_STEM = $(if $(SHEET),$(if $(wildcard examples/$(SHEET).csv),examples/$(SHEET).csv,),)
AUTO_CSV_INPUT = $(firstword $(CSV_FROM_SHEET_EXT) $(CSV_FROM_SHEET_STEM))
EFFECTIVE_INPUT = $(if $(AUTO_CSV_INPUT),$(AUTO_CSV_INPUT),$(INPUT))
EFFECTIVE_SHEET = $(if $(AUTO_CSV_INPUT),,$(SHEET))

SHEET_ARG = $(if $(EFFECTIVE_SHEET),--sheet $(EFFECTIVE_SHEET),)
TEMPLATE_ARG = $(if $(TEMPLATE),--template $(TEMPLATE),)

.PHONY: gen template clean help

gen:
	$(PYTHON) fsm_gen.py $(EFFECTIVE_INPUT) $(SHEET_ARG) $(TEMPLATE_ARG) -o $(OUT)

template:
	$(PYTHON) fsm_gen.py --new-adc-template examples/adc_controller_template.xlsx

clean:
	rm -rf generated generated_* sim_check csrc __pycache__ *.pyc
	find . -type d -name '__pycache__' -prune -exec rm -rf {} +
	rm -rf simv simv.daidir simv.vdb

help:
	@echo "make gen"
	@echo "make gen INPUT=examples/adc_controller_example.csv"
	@echo "make gen INPUT=examples/adc_controller.csv"
	@echo "make gen INPUT=examples/adc_controller_template.xlsx SHEET=adc_controller OUT=generated"
	@echo "make gen SHEET=traffic_light_example"
	@echo "make gen SHEET=traffic_light_example.csv"
	@echo "make template"
	@echo "make clean"
