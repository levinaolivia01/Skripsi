import sys
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor
import os

# Ambil path ke notebook dan file Excel dari argumen
notebook_path = sys.argv[1]
excel_path = sys.argv[2]

# Debug: Tampilkan path yang diterima
print(f"Path to notebook: {notebook_path}")
print(f"Path to Excel file: {excel_path}")

# Cek keberadaan file Excel
if not os.path.exists(excel_path):
    print(f"Error: File Excel '{excel_path}' tidak ditemukan.")
    sys.exit(1)

try:
    # Baca notebook
    with open(notebook_path) as f:
        notebook = nbformat.read(f, as_version=4)

    # Tambahkan variabel excel_path ke dalam sel pertama
    notebook.cells.insert(0, nbformat.v4.new_code_cell(f'excel_path = "{excel_path}"'))

    # Eksekusi notebook dengan Excel path sebagai variabel
    ep = ExecutePreprocessor(timeout=0, kernel_name='python3')
    ep.preprocess(notebook, {'metadata': {'path': './'}})

    # Tampilkan output dari notebook
    print("=== Output dari Jupyter Notebook ===")
    for cell in notebook['cells']:
        if cell['cell_type'] == 'code' and 'outputs' in cell:
            for output in cell['outputs']:
                if output.output_type == 'stream':
                    print(output.text)
                elif output.output_type == 'execute_result':
                    print(output['data']['text/plain'])

except Exception as e:
    print(f"Error saat mengeksekusi notebook: {e}")
