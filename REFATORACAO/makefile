.PHONY: all clean

# Nome do interpretador Python (pode variar dependendo do sistema)
PYTHON = python3

# Arquivos a serem compilados
FILES = main.py gui_functions.py aux_functions.py process_functions.py

# Nome do executável (opcional)
EXECUTABLE = main

# Regra padrão para compilar tudo
all: $(EXECUTABLE)

# Regra para compilar o executável
$(EXECUTABLE):
	$(PYTHON) -m py_compile $(FILES)

# Regra para limpar arquivos compilados
clean:
	rm -f *.pyc
