from function import *
from inserisci_entrata import inserisci_entrata
from inserisci_uscita import inserisci_uscita
from visualizza_dati import visualizza_dati

def main():
    print("1- Inserisci entrata.")
    print("2- Inserisci uscita.")
    print("3- Visualizza dati, saldo e salva in Excel.")
    choise = validate_input(1, 3)
    if choise == 1:
        inserisci_entrata()
    elif choise == 2:
        inserisci_uscita()
    elif choise == 3:
        visualizza_dati()

    choise = input("Vuoi continuare? (y/n) ").lower()
    if choise == "y" or choise == "s":
        main()
    else:
        print("Uscita dal programma.")
    
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nUscita dal programma.")
