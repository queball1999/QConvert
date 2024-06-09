import os
import sys
import subprocess
import mimetypes
from PySide6.QtWidgets import *
from PySide6.QtCore import *
from PySide6.QtGui import *

class ConversionThread(QThread):
    progress = Signal(int)
    finished = Signal(bool)
    error = Signal(str)
    output = Signal(str)

    def __init__(self, input_file, output_file, input_format, output_format, pdf_engine):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
        self.input_format = input_format
        self.output_format = output_format
        self.pdf_engine = pdf_engine

    def run(self):
        try:
            command = ['pandoc', '-f', self.input_format, '-t', self.output_format, '-o', self.output_file, self.input_file]
            
            if self.output_format == 'pdf':
                command.insert(1, f'--pdf-engine={self.pdf_engine}')
                
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.progress.emit(50)
            stdout, stderr = process.communicate()

            if process.returncode != 0:
                raise subprocess.CalledProcessError(process.returncode, command, output=stderr)
            
            self.progress.emit(100)
            self.output.emit(stdout.decode())
            self.finished.emit(True)
        except Exception as e:
            self.error.emit(str(e))
            self.finished.emit(False)


class FileConverter(QMainWindow):
    SUPPORTED_INPUT_FORMATS = ['epub', 'docx', 'md', 'html', 'txt']
    SUPPORTED_OUTPUT_FORMATS = ['pdf', 'docx', 'html', 'md', 'txt']
    SUPPORTED_PDF_ENGINES = ['pdflatex', 'xelatex', 'lualatex']

    def __init__(self):
        super().__init__()
        self.pdf_engine = 'xelatex'
        self.bulk_conversion = False
        self.initUI()

    def initUI(self):
        self.setWindowTitle('File Converter')
        self.setGeometry(300, 300, 600, 400)

        central_widget = QWidget()
        layout = QVBoxLayout()

        self.file_label = QLabel('No file selected', self)
        layout.addWidget(self.file_label)

        self.select_file_button = QPushButton('Select File', self)
        self.select_file_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_file_button)

        h_layout = QHBoxLayout()

        self.input_format_combo = QComboBox(self)
        self.input_format_combo.addItems(self.SUPPORTED_INPUT_FORMATS)
        self.input_format_combo.setEnabled(False)
        h_layout.addWidget(self.input_format_combo)

        arrow_label = QLabel('→', self)
        arrow_label.setAlignment(Qt.AlignCenter)
        h_layout.addWidget(arrow_label)

        self.output_format_combo = QComboBox(self)
        self.output_format_combo.addItems(self.SUPPORTED_OUTPUT_FORMATS)
        h_layout.addWidget(self.output_format_combo)

        layout.addLayout(h_layout)

        self.convert_button = QPushButton('Convert', self)
        self.convert_button.clicked.connect(self.convert_file)
        layout.addWidget(self.convert_button)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.output_text = QTextEdit(self)
        self.output_text.setReadOnly(True)
        self.output_text.setVisible(False)
        layout.addWidget(self.output_text)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.initMenu()

        if not self.check_pandoc_installed():
            self.show_pandoc_not_installed_message()
            sys.exit(1)
        elif not self.check_pdflatex_installed():
            self.show_pdflatex_not_installed_message()
            sys.exit(1)

    def initMenu(self):
        menubar = self.menuBar()
        options_menu = menubar.addMenu('Options')

        pdf_engine_menu = options_menu.addMenu('Select PDF Engine')
        for engine in self.SUPPORTED_PDF_ENGINES:
            action = QAction(engine, self, checkable=True)
            action.triggered.connect(lambda checked, e=engine: self.set_pdf_engine(e))
            pdf_engine_menu.addAction(action)
            if engine == self.pdf_engine:
                action.setChecked(True)

        bulk_conversion_action = QAction('Bulk Conversion by Directory', self, checkable=True)
        bulk_conversion_action.setChecked(self.bulk_conversion)
        bulk_conversion_action.triggered.connect(self.toggle_bulk_conversion)
        options_menu.addAction(bulk_conversion_action)

        display_output_action = QAction('Display Terminal Output', self, checkable=True)
        display_output_action.setChecked(False)
        display_output_action.triggered.connect(self.toggle_display_output)
        options_menu.addAction(display_output_action)

    def set_pdf_engine(self, engine):
        self.pdf_engine = engine
        print(f'Selected PDF engine: {self.pdf_engine}')

    def toggle_bulk_conversion(self, checked):
        self.bulk_conversion = checked
        if self.select_file_button.text() == 'Select File':
            self.select_file_button.setText('Select Folder')
        else:
            self.select_file_button.setText('Select File')
        print(f'Bulk conversion: {self.bulk_conversion}')

    def toggle_display_output(self, checked):
        self.output_text.setVisible(checked)
        print(f'Display terminal output: {checked}')

    def select_file(self):
        if self.bulk_conversion:
            self.input_directory = QFileDialog.getExistingDirectory(self, 'Select Directory')
            if self.input_directory:
                self.file_label.setText(self.input_directory)
            else:
                self.file_label.setText('No directory selected')
        else:
            file_filter = 'Supported files (*.epub *.docx *.md *.html *.txt);;All files (*.*)'
            self.input_file, _ = QFileDialog.getOpenFileName(self, 'Select File', '', file_filter)
            if self.input_file:
                self.file_label.setText(os.path.basename(self.input_file))
                self.detect_file_type()
            else:
                self.file_label.setText('No file selected')

    def detect_file_type(self):
        mime_type, _ = mimetypes.guess_type(self.input_file)
        if mime_type:
            file_extension = os.path.splitext(self.input_file)[1][1:].lower()
            if file_extension in self.SUPPORTED_INPUT_FORMATS:
                self.input_format_combo.setCurrentText(file_extension)
            else:
                QMessageBox.warning(self, 'Unsupported File', 'The selected file type is not supported.')
        else:
            QMessageBox.warning(self, 'File Type Detection Failed', 'Could not detect the file type.')

    def convert_file(self):
        if self.bulk_conversion:
            self.convert_bulk_files()
        else:
            self.convert_single_file()

    def convert_single_file(self):
        if not hasattr(self, 'input_file') or not self.input_file:
            QMessageBox.warning(self, 'Error', 'Please select a file to convert.')
            return
        
        parsed_input_file = self.input_file.split('.')[:-1]
        joined_input_file = '.'.join(parsed_input_file)
        input_format = self.input_format_combo.currentText()
        output_format = self.output_format_combo.currentText()
        output_file = QFileDialog.getSaveFileName(self, 'Save File', f'{joined_input_file}.{output_format}', f'*.{output_format}')[0]

        if not output_file:
            QMessageBox.warning(self, 'Error', 'Please specify an output file.')
            return

        self.progress_bar.setValue(0)
        self.thread = ConversionThread(self.input_file, output_file, input_format, output_format, self.pdf_engine)
        self.thread.progress.connect(self.update_progress)
        self.thread.finished.connect(self.on_conversion_finished)
        self.thread.error.connect(self.on_conversion_error)
        self.thread.output.connect(self.append_output)
        self.thread.start()

    def convert_bulk_files(self):
        if not hasattr(self, 'input_directory') or not self.input_directory:
            QMessageBox.warning(self, 'Error', 'Please select a file to convert.')
            return
        
        input_format = self.input_format_combo.currentText()
        output_format = self.output_format_combo.currentText()
        self.progress_bar.setValue(0)
        self.progress = 0
        for root, _, files in os.walk(self.input_directory):
            for file in files:
                if file.lower().endswith(f'.{input_format}'):
                    input_file = os.path.join(root, file)
                    parsed_input_file = input_file.split('.')[:-1]
                    joined_input_file = '.'.join(parsed_input_file)
                    output_file = os.path.join(root, f'{joined_input_file}.{output_format}')
                    #self.progress += 10
                    #self.progress_bar.setValue(self.progress)
                    self.thread = ConversionThread(input_file, output_file, input_format, output_format, self.pdf_engine)
                    self.thread.progress.connect(self.update_progress)
                    self.thread.finished.connect(self.on_conversion_finished)
                    self.thread.error.connect(self.on_conversion_error)
                    self.thread.output.connect(self.append_output)
                    self.thread.start()
                    self.thread.wait()  # Wait for the current file conversion to finish before starting the next
        self.progress_bar.setValue(100)
                    
    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def on_conversion_finished(self, success):
        if success:
            QMessageBox.information(self, 'Conversion Complete', 'File converted successfully.')
        else:
            QMessageBox.critical(self, 'Conversion Failed', 'An error occurred during conversion.')
        self.progress_bar.setValue(0)

    def on_conversion_error(self, error_message):
        QMessageBox.critical(self, 'Conversion Error', f'An error occurred during conversion:\n{error_message}')

    def append_output(self, output_message):
        self.output_text.append(output_message)

    def check_pdflatex_installed(self):
        try:
            result = subprocess.run(['pdflatex', '--version'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            print(f'pdflatex version check output: {result.stdout.decode()}')
            return True
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            print(f'Error checking pdflatex version: {str(e)}')
            print(f'System PATH: {os.environ["PATH"]}')
            return False
        
    def check_pandoc_installed(self):
        try:
            result = subprocess.run(['pandoc', '--version'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            print(f'Pandoc version check output: {result.stdout.decode()}')
            return True
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            print(f'Error checking Pandoc version: {str(e)}')
            print(f'System PATH: {os.environ["PATH"]}')
            return False

    def show_pandoc_not_installed_message(self):
        message = QMessageBox.critical(self, 'Pandoc Not Installed',
                             'Pandoc is not installed on your system. '
                             'Please install Pandoc from <a href="https://pandoc.org/installing.html">https://pandoc.org/installing.html</a>',
                             QMessageBox.Ok,
                             QMessageBox.Ok)
        
    def show_pdflatex_not_installed_message(self):
        message = QMessageBox.critical(self, 'pdflatex Not Installed',
                             'pdflatex is not installed on your system. '
                             'Please install a LaTeX distribution from <a href="https://miktex.org/download">https://miktex.org/download</a>',
                             QMessageBox.Ok,
                             QMessageBox.Ok)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FileConverter()
    ex.show()
    sys.exit(app.exec())