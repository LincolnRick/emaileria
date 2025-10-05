.PHONY: gui build-win build-mac

PYTHON ?= python
PYINSTALLER ?= pyinstaller

GUI_ENTRY := gui.py
SPEC_FILE := emaileria.spec

gui:
	$(PYTHON) $(GUI_ENTRY)

build-win:
	$(PYINSTALLER) --clean --onefile --noconsole $(SPEC_FILE)

build-mac:
	$(PYINSTALLER) --clean --onefile --windowed $(SPEC_FILE)
