import os
import asyncio
from classes.pandas_class import PandasHandler
from classes.playwright_class import PlayWriterHandler
from classes.openpyxl_class import OpenpyxlHandler
from classes.pypdf2_classes import PyPDF2_Handler
from classes.helpers import Helpers


pw = PlayWriterHandler()
ph = PandasHandler()
opxl = OpenpyxlHandler()
helper = Helpers()
pdf_handler = PyPDF2_Handler()

data_week1, data_week2 = asyncio.run(pw.run())


dataframe2 = ph.convert_to_dataframe(data_week2)
dataframe1 = ph.convert_to_dataframe(data_week1)
dataframe = ph.concat_dataframes([dataframe1, dataframe2])


opxl.set_data_in_excel(dataframe)
opxl.add_header()
opxl.format_heading_and_body_height()
opxl.format_width()
opxl.set_page_margin()
opxl.set_print_options()
opxl.add_formula_and_format_currency()
opxl.add_format_total_cell()
opxl.add_format_total_sum_cell()
opxl.add_format_amount_sum_cell()
opxl.add_format_declaration_cell()
opxl.add_format_signature_cell()
opxl.highlight_duplicate_cells()
opxl.highlight_anomalous_hours()


xl_name = helper.get_file_name(os.path.abspath(__file__), "excel") + ".xlsx"
opxl.save_to_excel(xl_name)
os.startfile(xl_name)


input("Press Enter to close Excel and generate PDF with password...")


pdf_name = helper.get_file_name(os.path.abspath(__file__), "pdf") + ".pdf"
helper.convert_excel_to_pdf(xl_name, pdf_name)
pdf_handler.set_pdf_password(pdf_name)

os.startfile(pdf_name)
