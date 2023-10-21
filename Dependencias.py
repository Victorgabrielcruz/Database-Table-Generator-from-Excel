import importlib

# Verifica se a biblioteca pandas está instalada
try:
    importlib.import_module('pandas')
    print("A biblioteca pandas está instalada.")
except ImportError:
    print("A biblioteca pandas não está instalada. Instalando...")
    import subprocess
    subprocess.run(['pip', 'install', 'pandas'])

# Verifica se a biblioteca pyodbc está instalada
try:
    importlib.import_module('pyodbc')
    print("A biblioteca pyodbc está instalada.")
except ImportError:
    print("A biblioteca pyodbc não está instalada. Instalando...")
    import subprocess
    subprocess.run(['pip', 'install', 'pyodbc'])
