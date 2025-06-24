import os
import tempfile
import uuid
import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET

from django.http import HttpResponse, FileResponse
from rest_framework import status
from rest_framework.parsers import MultiPartParser, JSONParser, FormParser
from rest_framework.response import Response
from rest_framework.views import APIView

import docx
from docx.shared import Pt
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from io import BytesIO


class DocConverterView(APIView):
    """
    API view for converting .doc/.docx files to PDF with variable replacement.
    
    Accepts a POST request with a document file and variable data.
    Returns a PDF file with variables replaced.
    """
    parser_classes = (MultiPartParser, FormParser, JSONParser)
    
    def post(self, request, *args, **kwargs):
        
        if 'file' not in request.FILES:
            return Response(
                {'error': 'No file provided'},
                status=status.HTTP_400_BAD_REQUEST
            )
        
        document_file = request.FILES['file']
        file_name = document_file.name
        file_extension = Path(file_name).suffix.lower()
        
        
        if file_extension not in ['.doc', '.docx']:
            return Response(
                {'error': 'Unsupported file format. Only .doc and .docx files are supported.'},
                status=status.HTTP_400_BAD_REQUEST
            )
        
      
        variables = {}
        if request.data.get('variables'):
            try:
                
                if isinstance(request.data['variables'], str):
                    try:
                        variables = json.loads(request.data['variables'])
                    except json.JSONDecodeError:
                        
                        var_name = request.data.get('variable_name', 'value')
                        variables = {var_name: request.data['variables']}
                elif isinstance(request.data['variables'], dict):
                    variables = request.data['variables']
                else:
                    return Response(
                        {'error': 'Variables must be a JSON object or string'},
                        status=status.HTTP_400_BAD_REQUEST
                    )
                
                
                print(f"Variables received: {variables}")
                
            except Exception as e:
                return Response(
                    {'error': f'Invalid variables format: {str(e)}'},
                    status=status.HTTP_400_BAD_REQUEST
                )
        
        
        replacement_mode = request.data.get('replacement_mode', 'dynamic').lower()
        
        try:
            
            converted_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'converted')
            os.makedirs(converted_dir, exist_ok=True)
            
           
            unique_id = str(uuid.uuid4())
            xml_path = os.path.join(converted_dir, f"{unique_id}.xml")
            pdf_path = os.path.join(converted_dir, f"{Path(file_name).stem}_converted_{unique_id}.pdf")
            
            
            if replacement_mode == 'static':
                modified_doc = self._replace_variables_static(document_file, variables)
            else:
                modified_doc = self._replace_variables_dynamic(document_file, variables)
            
            
            modified_doc_path = os.path.join(converted_dir, f"{Path(file_name).stem}_modified_{unique_id}.docx")
            modified_doc.save(modified_doc_path)
            print(f"Saved modified document to: {modified_doc_path}")
            
            
            xml_content = self._convert_doc_to_xml(modified_doc)
            
            
            with open(xml_path, 'wb') as xml_file:
                xml_file.write(xml_content)
            
            
            self._convert_xml_to_pdf(xml_content, pdf_path)
            
            
            response = FileResponse(
                open(pdf_path, 'rb'),
                content_type='application/pdf'
            )
            response['Content-Disposition'] = f'attachment; filename="{Path(file_name).stem}_converted.pdf"'
            return response
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            return Response(
                {'error': f'Conversion error: {str(e)}'},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )
    
    def _replace_variables_dynamic(self, doc_file, variables):
        """
        Replace variables in the document using Jinja-like placeholders {{variable_name}}.
        
        Args:
            doc_file: The uploaded document file
            variables (dict): Dictionary of variables to replace
            
        Returns:
            docx.Document: Modified document with variables replaced
        """
        try:
            
            document = docx.Document(doc_file)
            
            
            if not isinstance(variables, dict):
                variables = {}
            
            
            print(f"Variables for dynamic replacement: {variables}")
            
            
            for para in document.paragraphs:
                original_text = para.text
                modified_text = original_text
                
                
                for var_name, var_value in variables.items():
                    pattern = r'\{\{\s*' + re.escape(var_name) + r'\s*\}\}'
                    if re.search(pattern, modified_text):
                        print(f"Found '{var_name}' in paragraph: '{original_text}'")
                        modified_text = re.sub(pattern, str(var_value), modified_text)
                        print(f"After replacement: '{modified_text}'")
                
               
                if modified_text != original_text:
                   
                    for run in para.runs:
                        run.text = ""
                    
                    
                    para.add_run(modified_text)
            
            return document
            
        except Exception as e:
            raise Exception(f"Error replacing variables in document: {str(e)}")
    
    def _replace_variables_static(self, doc_file, variables):
        """
        Replace variables in the document using static sentence replacement.
        Looks for words that match variable names and replaces them.
        
        Args:
            doc_file: The uploaded document file
            variables (dict): Dictionary of variables to replace
            
        Returns:
            docx.Document: Modified document with variables replaced
        """
        try:
            
            document = docx.Document(doc_file)
            
            
            if not isinstance(variables, dict):
                variables = {}
            
            
            print(f"Variables for static replacement: {variables}")
            
            
            all_paragraphs = [p.text for p in document.paragraphs]
            print(f"Document paragraphs: {all_paragraphs}")
            
           
            for i, para in enumerate(document.paragraphs):
                original_text = para.text
                modified_text = original_text
                
                
                for var_name, var_value in variables.items():
                    
                    pattern = r'\b' + re.escape(var_name) + r'\b'
                    
                    
                    if re.search(pattern, modified_text, re.IGNORECASE):
                        print(f"Found word '{var_name}' in paragraph {i}: '{original_text}'")
                        
                        
                        modified_text = re.sub(pattern, str(var_value), modified_text, flags=re.IGNORECASE)
                        print(f"After replacement: '{modified_text}'")
                
               
                if modified_text != original_text:
                    para.clear()
                    para.add_run(modified_text)
            
            return document
            
        except Exception as e:
            raise Exception(f"Error replacing variables in document: {str(e)}")
    
    def _convert_doc_to_xml(self, document):
        """
        Convert document to XML format.
        
        Args:
            document: docx.Document object
            
        Returns:
            bytes: XML content as bytes
        """
        try:
            
            root = ET.Element("document")
            
            
            for para in document.paragraphs:
                p_elem = ET.SubElement(root, "paragraph")
                p_elem.text = para.text
                
                
                style = ET.SubElement(p_elem, "style")
                if para.style:
                    style.set("name", para.style.name if para.style.name else "")
                
                
                for run in para.runs:
                    run_elem = ET.SubElement(p_elem, "run")
                    run_elem.text = run.text
                    
                    if run.bold:
                        run_elem.set("bold", "true")
                    if run.italic:
                        run_elem.set("italic", "true")
                    if run.underline:
                        run_elem.set("underline", "true")
                    if run.font.size:
                        run_elem.set("size", str(run.font.size.pt))
                    if run.font.name:
                        run_elem.set("font", run.font.name)
            
            
            xml_data = ET.tostring(root, encoding='utf-8')
            return xml_data
            
        except Exception as e:
            raise Exception(f"Error converting document to XML: {str(e)}")
    
    def _convert_xml_to_pdf(self, xml_content, pdf_path):
        """
        Convert XML content to PDF using ReportLab (no HTML step).
        
        Args:
            xml_content (bytes): XML content with variables replaced
            pdf_path (str): Path to save the PDF file
            
        Returns:
            None
        """
        try:
            
            root = ET.fromstring(xml_content)
            
            
            doc = SimpleDocTemplate(
                pdf_path,
                pagesize=letter,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=72
            )
            
            
            styles = getSampleStyleSheet()
            normal_style = styles["Normal"]
            title_style = styles["Title"]
            heading_style = styles["Heading1"]
            
            
            bold_style = ParagraphStyle(
                'Bold', 
                parent=normal_style, 
                fontName='Helvetica-Bold'
            )
            
            italic_style = ParagraphStyle(
                'Italic', 
                parent=normal_style, 
                fontName='Helvetica-Oblique'
            )
            
            
            content = []
            
            
            for paragraph in root.findall('paragraph'):
                para_text = paragraph.text or ""
                
                
                style_elem = paragraph.find('style')
                para_style = normal_style
                
                if style_elem is not None:
                    style_name = style_elem.get('name', '')
                    if style_name and ('heading' in style_name.lower() or 'title' in style_name.lower()):
                        para_style = heading_style
                    elif style_name and 'title' in style_name.lower():
                        para_style = title_style
                
                
                run_texts = []
                for run in paragraph.findall('run'):
                    run_text = run.text or ""
                    
                    
                    if run.get('bold') == 'true':
                        run_text = f'<b>{run_text}</b>'
                    if run.get('italic') == 'true':
                        run_text = f'<i>{run_text}</i>'
                    if run.get('underline') == 'true':
                        run_text = f'<u>{run_text}</u>'
                    
                    run_texts.append(run_text)
                

                if run_texts:
                    para_text = "".join(run_texts)
                
                
                if para_text.strip():
                    content.append(Paragraph(para_text, para_style))
                    content.append(Spacer(1, 12))
            
            
            doc.build(content)
            
            return
            
        except Exception as e:
            raise Exception(f"Error converting XML to PDF: {str(e)}") 