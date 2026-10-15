import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QFormLayout, QInputDialog, QMessageBox
from PyQt5.QtCore import QDate
import xlsxwriter


class CalcoloScontiApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Scadenzario e Sconti")
        self.setGeometry(100, 100, 800, 600)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        file_label = QLabel("Percorso del file XLSX:", self)
        self.file_entry = QLineEdit(self)
        browse_button = QPushButton("Sfoglia", self)
        browse_button.clicked.connect(self.browse_file)

        layout.addWidget(file_label)
        layout.addWidget(self.file_entry)
        layout.addWidget(browse_button)

        form_layout = QFormLayout()

        self.percentuali_sconto = {}

        layout.addLayout(form_layout)
        crea_file_button = QPushButton("Crea File Scadenzario", self)
        crea_file_button.clicked.connect(self.crea_file_scadenzario)
        layout.addWidget(crea_file_button)

        calcola_button = QPushButton("Crea CSV sconti B2C e B2B", self)
        calcola_button.clicked.connect(self.calcola_sconti)

        layout.addWidget(calcola_button)

        rimuovi_sconti_button = QPushButton("Crea File Rimuovi Sconti", self)
        rimuovi_sconti_button.clicked.connect(self.rimuovi_sconti)
        layout.addWidget(rimuovi_sconti_button)

        central_widget.setLayout(layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Scegli un file XLSX", "", "Excel Files (*.xlsx);;All Files (*)")
        self.file_entry.setText(file_path)

    def calcola_sconti(self):
        try:
            file_path = self.file_entry.text()

            df = pd.read_excel(file_path)

            df['Scadenza'] = pd.to_datetime(df['Scadenza'], format="%d-%m-%Y")
            df = df.sort_values(by='Scadenza', ascending=True)

            # Estrai l'anno dalla colonna 'Scadenza'
            df['Anno'] = df['Scadenza'].dt.year

            first_three_months = df['Scadenza'].dt.month.unique()[:3]

            # Trova l'anno più piccolo tra i primi 3 mesi
            min_year = df[df['Scadenza'].dt.month.isin(first_three_months)]['Anno'].min()

            for mese in first_three_months:
                mese_testuale = QDate(2000, mese, 1).toString("MMMM")  # Ottieni il nome del mese in italiano
                sconto = float(self.show_sconto_dialog(mese_testuale))
                self.percentuali_sconto[mese] = sconto

            # Modifica: Seleziona solo i primi 3 mesi dell'anno più piccolo, considerando l'anno successivo se tra i mesi c'è dicembre
            df_first_three_months = df[((df['Scadenza'].dt.month.isin(first_three_months)) & (df['Anno'] == min_year)) | ((df['Scadenza'].dt.month >= 1) & (df['Scadenza'].dt.month <= 2) & (df['Anno'] == min_year + 1))].copy()


            for mese, sconto in self.percentuali_sconto.items():
                mask = df_first_three_months['Scadenza'].dt.month == mese
                df_first_three_months.loc[mask, 'Prezzo scontato Colonna 9'] = round(df_first_three_months.loc[mask, 'p_b2c'] * (1 - sconto / 100), 2)
                df_first_three_months.loc[mask, 'Prezzo scontato Colonna 10'] = round(df_first_three_months.loc[mask, 'p_b2b'] * df_first_three_months.loc[mask, 'QxCF'] * (1 - sconto / 100), 2)


            df_sconti_b2c = df_first_three_months[['Cd_AR', 'AR_Descrizione', 'Prezzo scontato Colonna 9']]
            df_sconti_b2c.columns = ['COD', 'Nome', 'Prezzo in offerta']

            df_sconti_b2b = df_first_three_months[['Cd_AR', 'AR_Descrizione', 'Prezzo scontato Colonna 10']]
            df_sconti_b2b.columns = ['COD', 'Nome', 'Prezzo in offerta']

            # Rimuovi la colonna 'Anno' se non è più necessaria
            df_first_three_months = df_first_three_months.drop(columns=['Anno'])

            df_sconti_b2c.to_csv('Sconti_B2C.csv', index=False, header=True)
            df_sconti_b2b.to_csv('Sconti_B2B.csv', index=False, header=True)

            print("File CSV creati con successo: Sconti_B2C.csv e Sconti_B2B.csv")

            self.show_success_popup("Operazione completata", "I file CSV sono stati creati con successo")


        except Exception as e:
            print(f"Si è verificato un errore: {str(e)}")

    def show_sconto_dialog(self, mese):
        sconto, okPressed = QInputDialog.getDouble(self, f"Inserisci la percentuale di sconto per {mese}", f"Percentuale di sconto per {mese}:", 0, 0, 100, 2)
        if okPressed:
            return sconto
        else:
            return 0

    def crea_file_scadenzario(self):
        file_path = self.file_entry.text()

        # Leggi il file Excel
        df = pd.read_excel(file_path)

        # Seleziona solo le colonne desiderate
        colonne_desiderate = ['Cd_AR', 'AR_Descrizione', 'Cd_ARMisura', 'Quantita', 'QuantitaDisp', 'p_b2c', 'p_b2b', 'QxCF', 'UltCost', 'Scadenza']
        df = df[colonne_desiderate]

        # Rinomina le colonne
        df.columns = ['Codice Articolo', 'Nome', 'UM', 'Quantita', 'QuantitaDisp', 'Prezzo B2C', 'Prezzo B2B', 'QxCF', 'UltCost', 'Scadenza']

        # Converte la colonna 'Scadenza' in formato datetime
        df['Scadenza'] = pd.to_datetime(df['Scadenza'], format="%d-%m-%Y")

        # Ordina il DataFrame per la colonna 'Scadenza' in ordine crescente (data più recente in cima)
        df = df.sort_values(by='Scadenza', ascending=True)

        # Formatta la colonna 'Scadenza' per mostrare solo la data senza l'orario
        df['Scadenza'] = df['Scadenza'].dt.strftime('%d/%m/%Y')

        # Crea un oggetto ExcelWriter utilizzando il motore xlsxwriter
        writer = pd.ExcelWriter('scadenzario.xlsx', engine='xlsxwriter')

        # Aggiungi bordi neri spessi alle celle con dati
        df = df.style.applymap(lambda x: "border: 4px solid black")

        # Scrivi il DataFrame nel foglio Excel
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Congela la riga 1
        worksheet = writer.sheets['Sheet1']
        worksheet.freeze_panes(1, 0)

        # Imposta la formattazione della colonna Scadenza come data in formato dd/mm/yyyy
        date_format = writer.book.add_format({'num_format': 'dd/mm/yyyy'})
        worksheet.set_column('J:J', None, date_format)

        # Salva il file Excel
        writer._save()

        print("File scadenzario.xlsx creato con successo.")

        self.show_success_popup("Operazione completata", "Il file scadenzario.csv è stato creato con successo.")

    def rimuovi_sconti(self):
        try:
            file_path = self.file_entry.text()
            df = pd.read_excel(file_path)

            # Crea un nuovo DataFrame con le colonne desiderate
            df_nuovo = df[['Cd_AR', 'AR_Descrizione']].copy()

            # Rinomina le colonne
            df_nuovo.columns = ['COD', 'Nome']

            # Crea una colonna "Prezzo in offerta" con valori vuoti
            df_nuovo['Prezzo in offerta'] = ''

            # Salva il nuovo DataFrame in un file CSV
            df_nuovo.to_csv('rimuovi_sconti.csv', index=False)

            print("File rimuovi_sconti.csv creato con successo.")

            self.show_success_popup("Operazione completata", "Il file rimuovi_sconti.csv è stato creato con successo.")

        except Exception as e:
            print(f"Si è verificato un errore: {str(e)}")

    def show_success_popup(self, title, message):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.exec_()


if __name__ == '__main__':
    app = QApplication([])
    window = CalcoloScontiApp()
    window.show()
    app.exec_()
