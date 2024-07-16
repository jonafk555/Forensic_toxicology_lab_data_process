import sys
import pandas as pd
import fitz  # PyMuPDF
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel

class DataCleaningApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Data Cleaning App')
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        self.label = QLabel("Please upload the PDF file that generate by WATERS's app.")
        layout.addWidget(self.label)

        self.btn = QPushButton('Upload PDF')
        self.btn.clicked.connect(self.uploadPDF)
        layout.addWidget(self.btn)

        self.setLayout(layout)

    def uploadPDF(self):
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "", "PDF Files (*.pdf);;All Files (*)", options=options)
        if filePath:
            self.label.setText(f'File loaded: {filePath}')
            self.processPDF(filePath)

    def processPDF(self, pdf_path):
        try:
            pdf_document = fitz.open(pdf_path)
            data = []
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text = page.get_text("text")
                lines = text.splitlines()
                for line in lines:
                    parts = line.split()
                    if len(parts) >= 8:
                        try:
                            ch = int(parts[0])
                            prnt_da = float(parts[1])
                            dau_da = float(parts[2])
                            dwell_s = float(parts[3])
                            cone_v = float(parts[4])
                            coll_ev = float(parts[5])
                            delay_s = parts[6]
                            compound = " ".join(parts[7:])
                            if "Intelli" in compound:
                               compound = compound.split("Intel")[0].strip()
                            formula = ""
                            new_compound = []
                            for part in compound.split():
                                if part.replace('.', '', 1).isdigit():
                                    formula = part
                                else:
                                    new_compound.append(part)
                            compound = " ".join(new_compound)
                            data.append([ch, prnt_da, dau_da, dwell_s, cone_v, coll_ev, delay_s, compound, formula])
                        except ValueError:
                            continue

            columns = ["Ch", "Prnt(Da)", "Dau(Da)", "Dwell(s)", "Cone(V)", "Coll(eV)", "Delay(s)", "Compound", "Formula"]
            df = pd.DataFrame(data, columns=columns)
            
            # Prompt user to select save location
            options = QFileDialog.Options()
            output_file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
            if output_file_path:
                df.to_excel(output_file_path, index=False)
                self.label.setText(f'Data processed and saved to {output_file_path}')
            else:
                self.label.setText('File save cancelled.')

        except Exception as e:
            self.label.setText(f'Error: {str(e)}')
            print(f'Error: {str(e)}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DataCleaningApp()
    ex.show()
    sys.exit(app.exec_())
