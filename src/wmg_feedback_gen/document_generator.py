import wmg_feedback_gen.core as core
from docxtpl import DocxTemplate
from docx.enum.text import WD_COLOR_INDEX
import openpyxl
import jinja2
from docx import Document
from io import BytesIO

def generate_doc(row_data: dict, 
                 template: DocxTemplate, 
                 output_filename: str,
                 jinja_env=None,
                 highlight=True):
    template.reset_replacements()
    template.render(row_data, jinja_env=jinja_env)

    # Save to an in-memory file
    template.save(output_filename)

    if highlight:
        docx = Document(output_filename)
        for table in docx.tables:
            for row in table.rows:
                if len(row.cells) != 9: # not the correct table
                    continue

                comments = row.cells[2]
                categories = ["OUTSTANDING", "DISTINCTION", "GOOD", "PASS", "MARGINAL", "FAIL"]
                lookup = dict(zip(categories, row.cells[3:9]))
                category = comments.text.strip().split()[-1]
                
                # Highlight the category in the document
                try:
                    for p in lookup[category].paragraphs:
                        for run in p.runs:
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                            print(f"Highlighting {category} in {output_filename}")
                except KeyError:
                    continue

        # Save the modified document to the output file
        docx.save(output_filename)


def generate( 
    xlsx_filename: str,
    template_filename: str,
    worksheet: str = "marks",
    output_filename: str = "feedback/feedback_{{STUDENTID}}.docx",
    validators: dict = core.default_validators,
    jinga_env=None,
    highlight=True):

    tpl = DocxTemplate(template_filename)

    if jinga_env is None:
        # Create a new Jinja2 environment if not provided
        jinja_env = jinja2.Environment()
        jinja_env.filters['mark_category'] = core.mark_category

    variables = tpl.get_undeclared_template_variables(jinja_env=jinja_env)

    workbook = openpyxl.load_workbook(xlsx_filename, data_only=True)

    for row_data in core.process_to_dicts(workbook[worksheet], variables, validators):
        print( row_data['STUDENTID'] )
        generate_doc(row_data,
                     tpl,
                     core.gen_filename(output_filename, row_data),
                     jinja_env=jinja_env,
                     highlight=highlight)

