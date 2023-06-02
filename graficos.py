# External imports
import sys
import json
from datetime import date
import xlsxwriter
import sidra_helpers


# Local imports
import scripts


def main(argv: list[str]):
    if len(argv) < 2:
        sys.exit("Too few arguments. Try 'python graficos.py help'.")

    if len(argv) < 3 and "help" in argv:
        print("Usage: python graficos.py <charts you want>")
        sys.exit()

    if "config" in argv:
        access_config(argv[1:])

    filename = handle_filename(argv)
    workbook = xlsxwriter.Workbook(f"files/{filename}")

    credits = [
        "Tabela feita automaticamente em Python. Código em:",
        "https://github.com/GuilhermeFrainer/graficos_excel",
        "",
        "Fontes dos dados:",
        "",
    ]

    function_dict = {
        'caged': scripts.caged,
        'cpi': scripts.cpi,
        'desemprego': scripts.desemprego,
        'dollar_euro': scripts.dollar_euro,
        'indice_vendas': scripts.indice_vendas,
        'ipca': scripts.ipca,
        'massa_rendimentos': scripts.massa_rendimentos,
        'pea': scripts.pea,
        'pib': scripts.pib,
        'treasury': scripts.treasury,
        'var_vendas': scripts.var_vendas,
    }
    for argument in argv[1:]:
        try:
            function_dict[argument](workbook, credits)
            print(f"Successfully created {argument} Excel sheet.")
        except KeyError:
            print(f"{argument} isn't an available script.")


    sidra_helpers.make_credits(workbook, credits)
    workbook.close()


# 'argv' here skips the name of the python program being run (i.e. it's shorter than usual)
def access_config(argv: list[str]):
    if len(argv) == 1:
        sys.exit("Please input the file whose config you wish to access.")
    elif len(argv) > 2:
        sys.exit("You may only access the config of one file at a time.")
    
    for arg in argv:
        if arg == "config":
            continue
        try:
            with open(f"config/{arg}.json", "r") as file:
                config_dict = json.load(file)
        except FileNotFoundError:
            sys.exit(f"{arg} is not an available file.")

    print(json.dumps(config_dict, indent=4))
    sys.exit()


# Determines the name of the file to be created
def handle_filename(argv: list[str]) -> str:
    if "-o" in argv:
        filename_index = argv.index("-o") + 1
        filename = argv[filename_index]
        if "." in filename:
            sys.exit("Invalid filename: please do not put an extension in the filename.")
        
        return f"{filename} {date.today().isoformat()}.xlsx"            

    else:
        return f"Gráficos {date.today().isoformat()}.xlsx"


if __name__=="__main__":
    main(sys.argv)

