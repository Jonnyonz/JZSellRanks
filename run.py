import sys
import os
import streamlit.web.cli as stcli

if __name__ == '__main__':
    # Obtener la ruta del archivo app.py
    script_path = os.path.join(os.path.dirname(__file__), 'app.py')
    sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
    sys.exit(stcli.main())