import pdfplumber
with pdfplumber.open('Statements/icici.pdf') as pdf:
    for i, strat in enumerate([None, {'vertical_strategy': 'lines_strict', 'horizontal_strategy': 'lines_strict'}, {'vertical_strategy': 'lines', 'horizontal_strategy': 'text'}, {'vertical_strategy': 'text', 'horizontal_strategy': 'text'}]):
        try:
            t = pdf.pages[0].extract_tables(table_settings=strat)
            print(f'Strat {i} n={len(t)}')
            if t: print(t[0][0])
        except Exception as e:
            print(f'Strat {i} Error: {e}')
