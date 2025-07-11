import wmg_feedback_gen.core as core
from docxtpl import DocxTemplate
import openpyxl

def generate_doc(row_data: dict, 
                 template: DocxTemplate, 
                 output_filename: str):
    template.reset_replacements()
    template.render(row_data)
    template.save(output_filename)

def generate( 
    xlsx_filename: str,
    template_filename: str,
    worksheet: str = "marks",
    output_filename: str = "feedback/feedback_{{STUDENTID}}.docx",
    validators: dict = core.default_validators):

    tpl = DocxTemplate(template_filename)
    variables = tpl.get_undeclared_template_variables()

    workbook = openpyxl.load_workbook(xlsx_filename, data_only=True)

    for row_data in core.process_to_dicts(workbook[worksheet], variables, validators):
        generate_doc(row_data,
                     tpl,
                     core.gen_filename(output_filename, row_data))
