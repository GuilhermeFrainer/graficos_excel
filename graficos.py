# External imports
import sys
import win32com.client

# Local imports
import scripts

def main(argv: list[str]):
    if len(argv) < 2:
        #
        # ADD USAGE MESSAGE
        #
        sys.exit("Too few arguments")

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


if __name__=="__main__":
    main(sys.argv)

