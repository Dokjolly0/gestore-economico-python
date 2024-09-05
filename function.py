import datetime
import sys
import openpyxl
from openpyxl.styles import PatternFill

def exit():
    sys.exit()

def validate_input(start: int, end: int, text = "Scegli un'opzione: ", exclude_value = None):
    if isinstance(start, int) and isinstance(end, int):
        while True:
            choise_input = input(text)
            try:
                choise = int(choise_input)
                if start <= choise <= end:
                    return choise
            except ValueError:
                if choise_input == exclude_value:
                    if isinstance(exclude_value, list):
                        for element in exclude_value:
                            if element == choise_input:
                                return element
                    return choise_input
                else:
                    continue
            except KeyboardInterrupt:
                print("\nUscita dal programma.")
                exit()
            except:
                continue
            
def validate_input_float(start: int, end: float, text = "Scegli un'opzione: "):
    if isinstance(start, int) and isinstance(end, float):
        while True:
            try:
                choise = float(input(text)) 
                if start <= choise <= end:
                    return choise
            except KeyboardInterrupt:
                print("\nUscita dal programma.")
                exit()
            except:
                continue

def get_date():
    data_input = input("Inserisci la data di accredito dell'importo o enter per la data attuale (gg/mm/aaaa): ").lower()
    if data_input == '':
        data = datetime.datetime.now()
        return data
    else:
        try:
            data = datetime.datetime.strptime(data_input, "%d/%m/%Y")
            return data
        except:
            print("Formato della data non corretto, reinseriscilo.")
            get_date()

def get_day():
    giorno_input = input("Inserisci il giorno della settimana o enter per il giorno attuale: ").lower()
    giorni = ["lunedi", "lunedì", "martedi", "martedì", "mercoledi", "mercoledì", "giovedi", "giovedì", "venerdi", "venerdì", "sabato", "domenica"]

    if giorno_input == '':
        index = datetime.datetime.now().weekday()
        return get_day_to_index(index)

    for giorno in giorni:
        if giorno_input == giorno:
            return giorno_input
    get_day()

def get_day_to_index(day):
    if isinstance(day, int):
        if day == 0:
            return "lunedì"
        elif day == 1:
            return "martedì"
        elif day == 2:
            return "mercoledì"
        elif day == 3:
            return "giovedì"
        elif day == 4:
            return "venerdì"
        elif day == 5:
            return "sabato"
        elif day == 6:
            return "domenica"
        else:
            return False
        
def ridimensiona_file_excel(sheet):
    # Ridimensiona automaticamente la larghezza delle colonne
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Ottieni la lettera della colonna (es. 'A')
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)  # Aggiungi un po' di spazio extra
        sheet.column_dimensions[column].width = adjusted_width
    return sheet


def estrai_colonna(sheet, index_colonna):
    # Inizializza una lista vuota per salvare i valori
    valori_colonna = []

    # Itera attraverso tutte le righe del foglio di lavoro per la colonna specificata
    for cell in sheet[index_colonna]:
        if cell.value is not None:  # Aggiungi solo valori non nulli
            valori_colonna.append(cell.value)
    return valori_colonna

def conta_righe_colonna(sheet, colonna_lettera):
    # Conta le celle non vuote nella colonna specificata
    conteggio = 0
    for cell in sheet[colonna_lettera]:
        if cell.value is not None:
            conteggio += 1
    return conteggio


def filter_date(date_list, mese: int, anno: int):
    try:
        # Verifica che mese e anno siano interi
        if (not isinstance(mese, int)) or (not isinstance(anno, int)):
            print(f"Mese: {mese} \nAnno: {anno}")
            print("Mese o anno non corretti.")
            return False

        # Converte le stringhe in oggetti datetime, se necessario
        date_objects = []
        for data in date_list:
            if isinstance(data, str):
                # Prova a convertire la stringa in un oggetto datetime con formato 'DD/MM/YYYY'
                try:
                    data = datetime.datetime.strptime(data, "%d/%m/%Y")  # Formato 'DD/MM/YYYY'
                except ValueError:
                    print(f"Formato data non valido: {data}")
                    continue
            date_objects.append(data)

        # print(f"Oggetti datetime convertiti: {date_objects}")
        # print(f"Filtriamo per mese: {mese}, anno: {anno}")

        # Filtra le date in base al mese e anno
        filtered_dates = [data for data in date_objects if data.month == mese and data.year == anno]

        #print(f"Date filtrate: {filtered_dates}")

        # Riconverte le date filtrate in stringhe formato 'DD/MM/YYYY'
        return [data.strftime("%d/%m/%Y") for data in filtered_dates]

    except Exception as e:
        print(f"Errore filtra date -> {str(e)}")



# date_list = ["05/08/2020", "12/09/2021", "05/08/2020", "01/01/2022", "05/09/2024"]
# print(filter_date(date_list, 9, 2024))  # Filtra per agosto 2020