# External imports
import sys
import json

# Local imports
import scripts


def main(argv: list[str]):
    if len(argv) < 2:
        sys.exit("Too few arguments.")

    if "config" in argv:
        access_config(argv[1:])

    function_dict = {
        'caged': scripts.caged,
        'cpi': scripts.cpi,
        'desemprego': scripts.desemprego,
        'dollar_euro': scripts.dollar_euro,
        'indice_vendas': scripts.indice_vendas,
        'indice_volume': scripts.indice_volume,
        'ipca': scripts.ipca,
        'massa_rendimentos': scripts.massa_rendimentos,
        'pea': scripts.pea,
        'pib': scripts.pib,
        'treasury': scripts.treasury,
    }
    for argument in argv[1:]:
        try:
            function_dict[argument]()
            print(f"Successfully created {argument} Excel file.")
        except KeyError:
            print(f"{argument} isn't an available script.")


# argv here skips the name of the python program being run (i.e. it's shorter than usual)
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


if __name__=="__main__":
    main(sys.argv)

