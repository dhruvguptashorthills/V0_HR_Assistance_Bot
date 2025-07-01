import base64
import io
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import fitz  # PyMuPDF
import copy
import re
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.shared import qn
from docx.enum.dml import MSO_THEME_COLOR_INDEX

class DocxUtils:
    @staticmethod
    def clean_na_values(data):
        """
        Recursively clean 'NA', 'N/A', empty strings, and None values from resume data.
        Returns a cleaned copy of the data.
        """
        if isinstance(data, dict):
            cleaned = {}
            for key, value in data.items():
                cleaned_value = DocxUtils.clean_na_values(value)
                if cleaned_value is not None and cleaned_value != '':
                    cleaned[key] = cleaned_value
            return cleaned
        elif isinstance(data, list):
            cleaned_list = []
            for item in data:
                cleaned_item = DocxUtils.clean_na_values(item)
                if cleaned_item is not None and cleaned_item != '':
                    if isinstance(cleaned_item, dict) and cleaned_item:
                        cleaned_list.append(cleaned_item)
                    elif not isinstance(cleaned_item, dict):
                        cleaned_list.append(cleaned_item)
            return cleaned_list
        elif isinstance(data, str):
            cleaned_str = data.strip()
            na_values = {'na', 'n/a', 'not applicable', 'not available', 'none', 'null', '-', ''}
            if cleaned_str.lower() in na_values:
                return None
            return cleaned_str
        else:
            return data

    @staticmethod
    def get_base64_image(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode("utf-8")

    @staticmethod
    def get_base64_pdf(pdf_file):
        pdf_file.seek(0)
        return base64.b64encode(pdf_file.read()).decode("utf-8")

    @staticmethod
    def clean_html_text(text):
        """Clean HTML tags and return text with bold formatting info"""
        if not text:
            return []
        
        text = str(text)
        bold_parts = []
        strong_pattern = r'<strong>(.*?)</strong>'
        matches = list(re.finditer(strong_pattern, text, re.IGNORECASE))
        
        if matches:
            current_pos = 0
            for match in matches:
                if match.start() > current_pos:
                    bold_parts.append((text[current_pos:match.start()], False))
                bold_parts.append((match.group(1), True))
                current_pos = match.end()
            if current_pos < len(text):
                bold_parts.append((text[current_pos:], False))
        else:
            bold_parts = [(text, False)]
        
        cleaned_parts = []
        for part_text, is_bold in bold_parts:
            clean_text = re.sub(r'<[^>]+>', '', part_text)
            if clean_text.strip():
                cleaned_parts.append((clean_text, is_bold))
        
        return cleaned_parts

    @staticmethod
    def add_formatted_text(paragraph, text_parts, font_size=10, font_color=RGBColor(34, 34, 34)):
        """Add text with mixed formatting to a paragraph"""
        for text, is_bold in text_parts:
            if text.strip():
                run = paragraph.add_run(text)
                run.font.name = 'Montserrat'
                run.font.size = Pt(font_size)
                run.font.color.rgb = font_color
                if is_bold:
                    run.bold = True

    @staticmethod
    def add_section_title(container, title, margin_top=18):
        """Add a section title with consistent formatting matching PDF template"""
        para = container.add_paragraph()
        para.paragraph_format.space_before = Pt(margin_top)
        para.paragraph_format.space_after = Pt(4)
        run = para.add_run(title.upper())
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(242, 93, 93)
        run.font.name = 'Montserrat'
        return para

    @staticmethod
    def add_triangle_bullet_point(container, text_parts, indent=22):
        """Add bullet point with triangle (▶) matching PDF template"""
        para = container.add_paragraph()
        para.paragraph_format.left_indent = Pt(indent)
        para.paragraph_format.space_after = Pt(6)
        
        triangle年后 = para.add_run("▶ ")
        triangle_run.font.name = 'Montserrat'
        triangle_run.font.size = Pt(13)
        triangle_run.font.color.rgb = RGBColor(242, 93, 93)
        triangle_run.bold = True
        
        DocxUtils.add_formatted_text(para, text_parts, font_size=12)
        return para

    @staticmethod
    def add_robust_page_border(doc):
        """Add page border with enhanced compatibility for Word web/desktop and PDF export"""
        try:
            for section in doc.sections:
                sectPr = section._sectPr
                pgBorders = OxmlElement('w:pgBorders')
                pgBorders.set(qn('w:offsetFrom'), 'page')
                pgBorders.set(qn('w:display'), 'allPages')
                
                for border_name in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{border_name}')
                    border.set(qn('w:val'), 'single')
                    border.set(qn('w:sz'), '6')
                    border.set(qn('w:space'), '0')
                    border.set(qn('w:color'), 'F25D5D')
                    pgBorders.append(border)
                
                sectPr.append(pgBorders)
        except Exception as e:
            print(f"Could not add page border: {e}")

    @staticmethod
    def create_compatible_table(parent, rows, cols, width_inches=None):
        """Create a table with enhanced compatibility for Word web/desktop"""
        try:
            if hasattr(parent, '_sectPr') or 'header' in str(type(parent)).lower() or 'footer' in str(type(parent)).lower():
                if width_inches is None:
                    width_inches = 8.0
                table = parent.add_table(rows=rows, cols=cols, width=Inches(width_inches))
            else:
                table = parent.add_table(rows=rows, cols=cols)
        except Exception as e:
            try:
                if width_inches is None:
                    width_inches = 8.0
                table = parent.add_table(rows=rows, cols=cols, width=Inches(width_inches))
            except:
                table = parent.add_table(rows, cols)
        
        try:
            table.autofit = False
            table.allow_autofit = False
        except:
            pass
        
        try:
            tbl = table._tbl
            tblPr = tbl.tblPr
            tblLayout = OxmlElement('w:tblLayout')
            tblLayout.set(qn('w:type'), 'fixed')
            tblPr.append(tblLayout)
        except:
            pass
        
        return table

    @staticmethod
    def add_column_border(cell, border_side='right', color='D3D3D3', width='6'):
        """Add a border to a specific side of a table cell with fallback compatibility"""
        try:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = tcPr.find(qn('w:tcBorders'))
            if tcBorders is None:
                tcBorders = OxmlElement('w:tcBorders')
                tcPr.append(tcBorders)
            
            border = OxmlElement(f'w:{border_side}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), width)
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), color)
            tcBorders.append(border)
        except Exception as e:
            print(f"Could not add cell border: {e}")

    @staticmethod
    def remove_all_table_borders(table):
        """Remove all table borders for seamless layout with better compatibility"""
        try:
            tbl = table._tbl
            tblPr = tbl.tblPr
            tblBorders = OxmlElement('w:tblBorders')
            for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = OxmlElement(f'w:{border_name}')
                border.set(qn('w:val'), 'nil')
                tblBorders.append(border)
            tblPr.append(tblBorders)
            
            for row in table.rows:
                for cell in row.cells:
                    try:
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        tcBorders = OxmlElement('w:tcBorders')
                        for border_name in ['top', 'left', 'bottom', 'right']:
                            border = OxmlElement(f'w:{border_name}')
                            border.set(qn('w:val'), 'nil')
                            tcBorders.append(border)
                        tcPr.append(tcBorders)
                    except:
                        continue
        except Exception as e:
            print(f"Could not remove table borders: {e}")

    @staticmethod
    def set_standard_spacing(paragraph, space_before_pt=0, space_after_pt=0):
        """Set paragraph spacing using standard methods for better compatibility"""
        if space_before_pt > 0:
            paragraph.paragraph_format.space_before = Pt(space_before_pt)
        if space_after_pt > 0:
            paragraph.paragraph_format.space_after = Pt(space_after_pt)

    @staticmethod
    def apply_standard_font(run, font_name='Montserrat', font_size_pt=10, is_bold=False, color_rgb=None):
        """Apply font formatting using standard methods for better compatibility"""
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        if is_bold:
            run.bold = True
        if color_rgb:
            run.font.color.rgb = color_rgb

    @staticmethod
    def add_cell_background_compatible(cell, color_hex='F2F2F2'):
        """Add cell background with enhanced compatibility"""
        try:
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), color_hex)
            cell._tc.get_or_add_tcPr().append(shading_elm)
            
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_mar = OxmlElement('w:tcMar')
            
            for margin in ['top', 'left', 'bottom', 'right']:
                mar_elem = OxmlElement(f'w:{margin}')
                mar_elem.set(qn('w:w'), '144')
                mar_elem.set(qn('w:type'), 'dxa')
                tc_mar.append(mar_elem)
            
            tc_pr.append(tc_mar)
            return True
        except Exception as e:
            print(f"Could not add cell background: {e}")
            return False

    @staticmethod
    def create_compatible_footer_table(footer, width_inches=8.5):
        """Create a footer table with enhanced compatibility"""
        try:
            footer_table = footer.add_table(rows=1, cols=1)
            footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            footer_table.autofit = False
            footer_table.columns[0].width = Inches(width_inches)
            DocxUtils.remove_all_table_borders(footer_table)
            return footer_table
        except Exception as e:
            print(f"Could not create footer table: {e}")
            return None

    @staticmethod
    def ensure_table_column_borders(table, column_index=0, border_color='CCCCCC'):
        """Ensure table has proper column borders for clear separation"""
        try:
            if column_index < len(table.columns) - 1:
                for row in table.rows:
                    cell = row.cells[column_index]
                    DocxUtils.add_column_border(cell, 'right', border_color, '8')
        except Exception as e:
            print(f"Could not add column borders: {e}")

    @staticmethod
    def set_fixed_column_widths(table, left_width_inches, right_width_inches):
        """Set fixed column widths that will be maintained during PDF export"""
        try:
            table.columns[0].width = Inches(left_width_inches)
            table.columns[1].width = Inches(right_width_inches)
            
            tbl = table._tbl
            tblPr = tbl.tblPr
            existing_layout = tblPr.find(qn('w:tblLayout'))
            if existing_layout is not None:
                tblPr.remove(existing_layout)
            
            tblLayout = OxmlElement('w:tblLayout')
            tblLayout.set(qn('w:type'), 'fixed')
            tblPr.append(tblLayout)
            
            tblW = OxmlElement('w:tblW')
            tblW.set(qn('w:w'), str(int((left_width_inches + right_width_inches) * 1440)))
            tblW.set(qn('w:type'), 'dxa')
            tblPr.append(tblW)
        except Exception as e:
            print(f"Could not set fixed column widths: {e}")

    @staticmethod
    def lock_all_table_layouts(doc):
        """Lock all table layouts to fixed for consistent PDF export"""
        try:
            for table in doc.tables:
                try:
                    tbl = table._tbl
                    tblPr = tbl.tblPr
                    existing_layout = tblPr.find(qn('w:tblLayout'))
                    if existing_layout is not None:
                        tblPr.remove(existing_layout)
                    tblLayout = OxmlElement('w:tblLayout')
                    tblLayout.set(qn('w:type'), 'fixed')
                    tblPr.append(tblLayout)
                except:
                    continue
            
            for section in doc.sections:
                try:
                    if section.header:
                        for table in section.header.tables:
                            try:
                                tbl = table._tbl
                                tblPr = tbl.tblPr
                                existing_layout = tblPr.find(qn('w:tblLayout'))
                                if existing_layout is not None:
                                    tblPr.remove(existing_layout)
                                tblLayout = OxmlElement('w:tblLayout')
                                tblLayout.set(qn('w:type'), 'fixed')
                                tblPr.append(tblLayout)
                            except:
                                continue
                    if section.footer:
                        for table in section.footer.tables:
                            try:
                                tbl = table._tbl
                                tblPr = tbl.tblPr
                                existing_layout = tblPr.find(qn('w:tblLayout'))
                                if existing_layout is not None:
                                    tblPr.remove(existing_layout)
                                tblLayout = OxmlElement('w:tblLayout')
                                tblLayout.set(qn('w:type'), 'fixed')
                                tblPr.append(tblLayout)
                            except:
                                continue
                except:
                    continue
        except Exception as e:
            print(f"Could not lock table layouts: {e}")

    @staticmethod
    def optimize_for_pdf_export(doc):
        """Apply optimizations for better PDF export from Word"""
        try:
            for section in doc.sections:
                section.different_first_page_header_footer = False
                section.start_type = WD_SECTION.NEW_PAGE
                section.page_width = Inches(8.5)
                section.page_height = Inches(11)
        except Exception as e:
            print(f"Could not optimize for PDF export: {e}")

    @staticmethod
    def add_document_background(doc, bg_image_path="templates/bg.png"):
        """Add a full-page background image to the document with 17% opacity"""
        try:
            section = doc.sections[0]
            sectPr = section._sectPr
            wrap = OxmlElement('w:background')
            wrap.set(qn('w:color'), 'FFFFFF')  # White background fallback
            vml = OxmlElement('v:background')
            vml.set(qn('id'), 'bgImage')
            fill = OxmlElement('v:fill')
            fill.set(qn('type'), 'tile')  # Tile to cover entire page
            fill.set(qn('src'), bg_image_path)
            fill.set(qn('opacity'), '0.17')  # 17% opacity to match original watermark
            vml.append(fill)
            wrap.append(vml)
            sectPr.append(wrap)
            return True
        except Exception as e:
            print(f"Failed to add document background: {e}")
            import traceback
            traceback.print_exc()
            return False

    @staticmethod
    def add_grey_background(cell):
        """Add grey background to cell - compatibility wrapper"""
        return DocxUtils.add_cell_background_compatible(cell, 'F2F2F2')

    @staticmethod
    def remove_table_borders(table):
        """Remove table borders - compatibility wrapper"""
        return DocxUtils.remove_all_table_borders(table)

    @staticmethod
    def add_page_border(doc):
        """Add page border - compatibility wrapper"""
        return DocxUtils.add_robust_page_border(doc)

    @staticmethod
    def optimize_table_for_word(table):
        """Optimize table for Word - compatibility wrapper"""
        return DocxUtils.remove_all_table_borders(table)

    @staticmethod
    def add_word_optimized_spacing(paragraph, space_before=0, space_after=0, line_spacing=1.0):
        """Add spacing - compatibility wrapper"""
        return DocxUtils.set_standard_spacing(paragraph, space_before, space_after)

    @staticmethod
    def add_word_font_optimization(run, font_name='Montserrat', font_size=10, is_bold=False, color_rgb=None):
        """Font optimization - compatibility wrapper"""
        return DocxUtils.apply_standard_font(run, font_name, font_size, is_bold, color_rgb)

    @staticmethod
    def generate_docx(data, keywords=None, left_logo_path="templates/left_logo_small.png", right_logo_path="templates/right_logo_small.png"):
        """
        Generate a .docx resume matching the PDF template exactly with a background image.
        Returns a BytesIO object containing the Word file.
        """
        data_copy = copy.deepcopy(data)
        data_copy = DocxUtils.clean_na_values(data_copy)
        
        if data_copy.get('skills') and len(data_copy['skills']) > 18:
            data_copy['skills'] = data_copy['skills'][:18]
        
        if data_copy.get('certifications') and len(data_copy['certifications']) > 5:
            data_copy['certifications'] = data_copy['certifications'][:5]
        
        doc = docx.Document()
        
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(0.2)
            section.bottom_margin = Inches(0.2)
            section.left_margin = Inches(0.2)
            section.right_margin = Inches(0.2)
            section.header_distance = Inches(0.15)
            section.footer_distance = Inches(0.15)
            section.page_width = Inches(8.5)
            section.page_height = Inches(11)

        # Add document background instead of watermark
        DocxUtils.add_document_background(doc, bg_image_path="templates/bg.png")
        DocxUtils.add_robust_page_border(doc)
        
        header = doc.sections[0].header
        header.is_linked_to_previous = False
        
        for para in header.paragraphs:
            try:
                p = para._element
                p.getparent().remove(p)
            except:
                para.clear()
        
        header_table = DocxUtils.create_compatible_table(header, rows=1, cols=3, width_inches=8.1)
        header_table.columns[0].width = Inches(2.7)
        header_table.columns[1].width = Inches(2.7)
        header_table.columns[2].width = Inches(2.7)
        
        left_cell = header_table.cell(0, 0)
        left_para = left_cell.paragraphs[0]
        left_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        DocxUtils.add_word_optimized_spacing(left_para, space_before=2, space_after=0)
        left_para.paragraph_format.left_indent = Pt(12)
        try:
            left_run = left_para.add_run()
            left_run.add_picture(left_logo_path, height=Inches(0.35))
        except Exception:
            left_run = left_para.add_run("ShorthillsAI")
            DocxUtils.add_word_font_optimization(left_run, 'Montserrat', 10, True, RGBColor(242, 93, 93))
        
        right_cell = header_table.cell(0, 2)
        right_para = right_cell.paragraphs[0]
        right_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        DocxUtils.add_word_optimized_spacing(right_para, space_before=2, space_after=0)
        right_para.paragraph_format.right_indent = Pt(12)
        try:
            right_run = right_para.add_run()
            right_run.add_picture(right_logo_path, height=Inches(0.45))
        except Exception:
            right_run = right_para.add_run("Microsoft Partner")
            right_run.font.name = 'Montserrat'
            right_run.font.size = Pt(10)
            right_run.font.color.rgb = RGBColor(102, 102, 102)
        
        main_table = DocxUtils.create_compatible_table(doc, rows=1, cols=2, width_inches=8.1)
        DocxUtils.set_fixed_column_widths(main_table, 2.8, 5.3)
        DocxUtils.ensure_table_column_borders(main_table, 0, 'CCCCCC')
        
        left_cell = main_table.cell(0, 0)
        right_cell = main_table.cell(0, 1)
        
        left_cell._tc.clear_content()
        right_cell._tc.clear_content()

        left_cell_padding = left_cell.add_paragraph()
        left_cell_padding.paragraph_format.left_indent = Pt(12)

        name_title_table = DocxUtils.create_compatible_table(left_cell, rows=2, cols=1, width_inches=2.8)
        name_title_table.columns[0].width = Inches(2.8)
        
        name_cell = name_title_table.cell(0, 0)
        DocxUtils.add_grey_background(name_cell)
        name_para = name_cell.paragraphs[0]
        DocxUtils.add_word_optimized_spacing(name_para, space_before=0, space_after=2)
        name_para.paragraph_format.left_indent = Pt(8)
        name_run = name_para.add_run(data_copy.get('name', ''))
        DocxUtils.add_word_font_optimization(name_run, 'Montserrat', 17, True, RGBColor(242, 93, 93))
        
        title_cell = name_title_table.cell(1, 0)
        DocxUtils.add_grey_background(title_cell)
        title_para = title_cell.paragraphs[0]
        DocxUtils.add_word_optimized_spacing(title_para, space_after=10)
        title_para.paragraph_format.left_indent = Pt(8)
        title_parts = DocxUtils.clean_html_text(data_copy.get('title', ''))
        for text, is_bold in title_parts:
            if text.strip():
                title_run = title_para.add_run(text)
                DocxUtils.add_word_font_optimization(title_run, 'Montserrat', 13, True, RGBColor(34, 34, 34))

        if data_copy.get('skills'):
            skills_para = left_cell.add_paragraph()
            DocxUtils.add_word_optimized_spacing(skills_para, space_before=10, space_after=2)
            skills_para.paragraph_format.left_indent = Pt(12)
            skills_run = skills_para.add_run('SKILLS')
            DocxUtils.add_word_font_optimization(skills_run, 'Montserrat', 11, True, RGBColor(242, 93, 93))
            
            for skill in data_copy['skills']:
                skill_parts = DocxUtils.clean_html_text(skill)
                skill_para = left_cell.add_paragraph()
                DocxUtils.add_word_optimized_spacing(skill_para, space_after=3)
                skill_para.paragraph_format.left_indent = Pt(28)
                
                arrow_run = skill_para.add_run("▶ ")
                DocxUtils.add_word_font_optimization(arrow_run, 'Montserrat', 11, True, RGBColor(242, 93, 93))
                
                for text, is_bold in skill_parts:
                    if text.strip():
                        skill_run = skill_para.add_run(text)
                        DocxUtils.add_word_font_optimization(skill_run, 'Montserrat', 10, is_bold, RGBColor(34, 34, 34))

        # Education
        if data_copy.get('education'):
            edu_title_para = doc.add_paragraph()
            edu_title_para.paragraph_format.left_indent = Pt(12)
            edu_title_para.paragraph_format.space_before = Pt(18)
            edu_title_para.paragraph_format.space_after = Pt(4)
            edu_title_run = edu_title_para.add_run('EDUCATION')
            edu_title_run.bold = True
            edu_title_run.font.size = Pt(13)
            edu_title_run.font.color.rgb = RGBColor(242, 93, 93)
            edu_title_run.font.name = 'Montserrat'
            
            for edu in data_copy['education']:
                if isinstance(edu, dict):
                    para = doc.add_paragraph()
                    para.paragraph_format.left_indent = Pt(34)  # 12 + 22
                    para.paragraph_format.space_after = Pt(6)
                
                    arrow_run = para.add_run("▶ ")
                    arrow_run.font.name = 'Montserrat'
                    arrow_run.font.size = Pt(13)
                    arrow_run.font.color.rgb = RGBColor(242, 93, 93)
                    arrow_run.bold = True
                    
                    if edu.get('degree'):
                        degree_parts = DocxUtils.clean_html_text(edu['degree'])
                        DocxUtils.add_formatted_text(para, degree_parts, font_size=12)
                    
                    if edu.get('institution'):
                        para.add_run('\n')
                        inst_parts = DocxUtils.clean_html_text(edu['institution'])
                        for text, is_bold in inst_parts:
                            if text.strip():
                                inst_run = para.add_run(text)
                                inst_run.font.name = 'Montserrat'
                                inst_run.font.size = Pt(12)
                                inst_run.font.color.rgb = RGBColor(34, 34, 34)
                                inst_run.bold = True
                else:
                    edu_para = doc.add_paragraph()
                    edu_para.paragraph_format.left_indent = Pt(34)
                    edu_para.paragraph_format.space_after = Pt(6)
                    
                    arrow_run = edu_para.add_run("▶ ")
                    arrow_run.font.name = 'Montserrat'
                    arrow_run.font.size = Pt(13)
                    arrow_run.font.color.rgb = RGBColor(242, 93, 93)
                    arrow_run.bold = True
                    
                    edu_parts = DocxUtils.clean_html_text(str(edu))
                    DocxUtils.add_formatted_text(edu_para, edu_parts, font_size=12)

        # Certifications
        if data_copy.get('certifications'):
            cert_title_para = doc.add_paragraph()
            cert_title_para.paragraph_format.left_indent = Pt(12)
            cert_title_para.paragraph_format.space_before = Pt(18)
            cert_title_para.paragraph_format.space_after = Pt(4)
            cert_title_run = cert_title_para.add_run('CERTIFICATIONS')
            cert_title_run.bold = True
            cert_title_run.font.size = Pt(13)
            cert_title_run.font.color.rgb = RGBColor(242, 93, 93)
            cert_title_run.font.name = 'Montserrat'
            
            for cert in data_copy['certifications']:
                para = doc.add_paragraph()
                para.paragraph_format.left_indent = Pt(34)  # 12 + 22
                para.paragraph_format.space_after = Pt(6)
                
                arrow_run = para.add_run("▶ ")
                arrow_run.font.name = 'Montserrat'
                arrow_run.font.size = Pt(13)
                arrow_run.font.color.rgb = RGBColor(242, 93, 93)
                arrow_run.bold = True
                
                if isinstance(cert, dict):
                    # Track if we have any content to add
                    has_content = False
                    
                    if cert.get('title'):
                        title_parts = DocxUtils.clean_html_text(cert['title'])
                        DocxUtils.add_formatted_text(para, title_parts, font_size=12)
                        has_content = True
                    
                    # Add issuer (only if we have title or other content)
                    if cert.get('issuer') and has_content:
                        para.add_run('\n')
                        issuer_parts = DocxUtils.clean_html_text(cert['issuer'])
                        for text, is_bold in issuer_parts:
                            if text.strip():
                                issuer_run = para.add_run(text)
                                issuer_run.font.name = 'Montserrat'
                                issuer_run.font.size = Pt(12)
                                issuer_run.font.color.rgb = RGBColor(34, 34, 34)
                                issuer_run.bold = True
                    elif cert.get('issuer') and not has_content:
                        # If no title but have issuer, add issuer as main content
                        issuer_parts = DocxUtils.clean_html_text(cert['issuer'])
                        for text, is_bold in issuer_parts:
                            if text.strip():
                                issuer_run = para.add_run(text)
                                issuer_run.font.name = 'Montserrat'
                                issuer_run.font.size = Pt(12)
                                issuer_run.font.color.rgb = RGBColor(34, 34, 34)
                                issuer_run.bold = True
                                has_content = True
                    
                    # Add year (only if we have other content)
                    if cert.get('year') and has_content:
                        year_run = para.add_run(f"\n{cert['year']}")
                        year_run.font.name = 'Montserrat'
                        year_run.font.size = Pt(12)
                        year_run.font.color.rgb = RGBColor(34, 34, 34)
                else:
                    cert_parts = DocxUtils.clean_html_text(str(cert))
                    DocxUtils.add_formatted_text(para, cert_parts, font_size=12)

        # Footer with compatible design
        footer = doc.sections[0].footer
        
        # Clear any default footer content
        for para in footer.paragraphs:
            para.clear()
            
        # Create footer table for consistent background
        footer_table = DocxUtils.create_compatible_footer_table(footer, 8.0)
        if footer_table is not None:
            footer_cell = footer_table.cell(0, 0)
            DocxUtils.add_cell_background_compatible(footer_cell, 'F25D5D')
            
            footer_para = footer_cell.paragraphs[0]
            footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            DocxUtils.set_standard_spacing(footer_para, space_before_pt=6, space_after_pt=6)
            
            footer_run = footer_para.add_run("© www.shorthills.ai")
            DocxUtils.apply_standard_font(footer_run, 'Montserrat', 10, False, RGBColor(255, 255, 255))

        # Save document
        docx_file = io.BytesIO()
        doc.save(docx_file)
        docx_file.seek(0)
        return docx_file