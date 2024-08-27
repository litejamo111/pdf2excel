import sys
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, QLabel
import tabula
import pandas as pd

class PDFToExcelConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel('Select a PDF file to convert to Excel')
        layout.addWidget(self.label)

        self.btn = QPushButton('Select PDF File', self)
        self.btn.clicked.connect(self.showDialog)
        layout.addWidget(self.btn)

        self.setLayout(layout)
        self.setGeometry(300, 300, 300, 150)
        self.setWindowTitle('PDF to Excel Converter')
        self.show()

    def showDialog(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'Open file', '/home', "PDF files (*.pdf)")
        if fname:
            self.convert_pdf_to_excel(fname)

    def convert_pdf_to_excel(self, pdf_path):
        try:
            # Read PDF file
            tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
            
            # Combine all tables into a single DataFrame
            df = pd.concat(tables, ignore_index=True)
            
            # Save to Excel
            excel_path = pdf_path.rsplit('.', 1)[0] + '.xlsx'
            df.to_excel(excel_path, index=False)
            
            self.label.setText(f'Successfully converted to {excel_path}')
        except Exception as e:
            self.label.setText(f'Error: {str(e)}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFToExcelConverter()
    sys.exit(app.exec())