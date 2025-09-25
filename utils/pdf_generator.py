"""
PDF Generation utilities for PWD Tools
Generate PDF documents for bills, reports, and certificates
"""

try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

import io
from datetime import datetime
import pandas as pd

class PDFGenerator:
    """PDF generation utility class"""
    
    def __init__(self):
        """Initialize PDF generator"""
        if not REPORTLAB_AVAILABLE:
            raise ImportError("ReportLab is required for PDF generation. Install with: pip install reportlab")
        
        self.styles = getSampleStyleSheet()
        self.setup_custom_styles()
    
    def setup_custom_styles(self):
        """Setup custom styles for PWD documents"""
        # PWD Header style
        self.styles.add(ParagraphStyle(
            name='PWDHeader',
            parent=self.styles['Heading1'],
            fontSize=18,
            textColor=colors.HexColor('#004E89'),
            alignment=TA_CENTER,
            spaceAfter=12
        ))
        
        # PWD Subheader style
        self.styles.add(ParagraphStyle(
            name='PWDSubHeader',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=colors.HexColor('#FF6B35'),
            alignment=TA_CENTER,
            spaceAfter=8
        ))
        
        # PWD Normal style
        self.styles.add(ParagraphStyle(
            name='PWDNormal',
            parent=self.styles['Normal'],
            fontSize=11,
            spaceAfter=6
        ))
        
        # PWD Table Header style
        self.styles.add(ParagraphStyle(
            name='PWDTableHeader',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=colors.white,
            alignment=TA_CENTER
        ))
    
    def create_bill_pdf(self, bill_data, output_path=None):
        """Generate PDF for bill document"""
        if output_path is None:
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
        else:
            doc = SimpleDocTemplate(output_path, pagesize=A4)
        
        # Build content
        story = []
        
        # Header
        header = Paragraph("राजस्थान सरकार<br/>लोक निर्माण विभाग<br/>Government of Rajasthan<br/>Public Works Department", self.styles['PWDHeader'])
        story.append(header)
        story.append(Spacer(1, 12))
        
        # Bill title
        bill_title = Paragraph(f"BILL NO: {bill_data.get('bill_number', 'N/A')}", self.styles['PWDSubHeader'])
        story.append(bill_title)
        story.append(Spacer(1, 12))
        
        # Bill details table
        bill_details = [
            ['Bill Date:', bill_data.get('bill_date', 'N/A')],
            ['Contractor Name:', bill_data.get('contractor_name', 'N/A')],
            ['Project Name:', bill_data.get('project_name', 'N/A')],
            ['Work Order No.:', bill_data.get('work_order_no', 'N/A')],
            ['Agreement Amount:', f"₹{bill_data.get('agreement_amount', 0):,.2f}"],
            ['Bill Amount:', f"₹{bill_data.get('bill_amount', 0):,.2f}"]
        ]
        
        details_table = Table(bill_details, colWidths=[2*inch, 4*inch])
        details_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F0F2F6')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(details_table)
        story.append(Spacer(1, 12))
        
        # Work description
        if bill_data.get('work_description'):
            work_desc = Paragraph(f"<b>Work Description:</b><br/>{bill_data['work_description']}", self.styles['PWDNormal'])
            story.append(work_desc)
            story.append(Spacer(1, 12))
        
        # Bill items table
        if 'items' in bill_data and bill_data['items']:
            items_header = Paragraph("Bill Items", self.styles['PWDSubHeader'])
            story.append(items_header)
            
            # Create items table
            items_data = [['S.No.', 'Description', 'Unit', 'Quantity', 'Rate (₹)', 'Amount (₹)']]
            
            for i, item in enumerate(bill_data['items'], 1):
                items_data.append([
                    str(i),
                    item.get('description', ''),
                    item.get('unit', ''),
                    f"{item.get('quantity', 0):.2f}",
                    f"{item.get('rate', 0):,.2f}",
                    f"{item.get('total', 0):,.2f}"
                ])
            
            items_table = Table(items_data, colWidths=[0.5*inch, 2.5*inch, 0.8*inch, 0.8*inch, 1*inch, 1*inch])
            items_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#004E89')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            story.append(items_table)
            story.append(Spacer(1, 12))
        
        # Deductions table
        if 'deductions' in bill_data and bill_data['deductions']:
            deductions_header = Paragraph("Deductions", self.styles['PWDSubHeader'])
            story.append(deductions_header)
            
            deductions_data = [['Deduction Type', 'Rate (%)', 'Amount (₹)']]
            total_deductions = 0
            
            for deduction in bill_data['deductions']:
                deductions_data.append([
                    deduction.get('type', ''),
                    f"{deduction.get('rate', 0):.2f}%" if deduction.get('rate', 0) > 0 else 'Fixed',
                    f"{deduction.get('amount', 0):,.2f}"
                ])
                total_deductions += deduction.get('amount', 0)
            
            deductions_data.append(['Total Deductions', '', f"{total_deductions:,.2f}"])
            
            deductions_table = Table(deductions_data, colWidths=[3*inch, 1.5*inch, 2*inch])
            deductions_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FF6B35')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F0F2F6')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('ALIGN', (0, 1), (0, -2), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            story.append(deductions_table)
            story.append(Spacer(1, 12))
        
        # Summary
        net_amount = bill_data.get('bill_amount', 0) - sum([d.get('amount', 0) for d in bill_data.get('deductions', [])])
        
        summary_data = [
            ['Bill Amount:', f"₹{bill_data.get('bill_amount', 0):,.2f}"],
            ['Total Deductions:', f"₹{sum([d.get('amount', 0) for d in bill_data.get('deductions', [])]):,.2f}"],
            ['Net Payable Amount:', f"₹{net_amount:,.2f}"]
        ]
        
        summary_table = Table(summary_data, colWidths=[3*inch, 2*inch])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#1A8A16')),
            ('TEXTCOLOR', (0, -1), (-1, -1), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -2), 'Helvetica'),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(summary_table)
        story.append(Spacer(1, 20))
        
        # Signature section
        signature_data = [
            ['Contractor Signature', 'Assistant Engineer', 'Executive Engineer'],
            ['_' * 20, '_' * 20, '_' * 20]
        ]
        
        signature_table = Table(signature_data, colWidths=[2.2*inch, 2.2*inch, 2.2*inch])
        signature_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6)
        ]))
        
        story.append(signature_table)
        
        # Footer
        footer = Paragraph(
            f"Generated on: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br/>PWD Tools - Bill Generator",
            self.styles['Normal']
        )
        story.append(Spacer(1, 20))
        story.append(footer)
        
        # Build PDF
        doc.build(story)
        
        if output_path is None:
            buffer.seek(0)
            return buffer.getvalue()
        
        return output_path
    
    def create_emd_refund_pdf(self, emd_data, output_path=None):
        """Generate PDF for EMD refund certificate"""
        if output_path is None:
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
        else:
            doc = SimpleDocTemplate(output_path, pagesize=A4)
        
        story = []
        
        # Header
        header = Paragraph("राजस्थान सरकार<br/>लोक निर्माण विभाग<br/>EMD REFUND CERTIFICATE", self.styles['PWDHeader'])
        story.append(header)
        story.append(Spacer(1, 20))
        
        # EMD details
        emd_details = [
            ['Tender Number:', emd_data.get('tender_number', 'N/A')],
            ['Contractor/Bidder Name:', emd_data.get('contractor_name', 'N/A')],
            ['EMD Amount:', f"₹{emd_data.get('emd_amount', 0):,.2f}"],
            ['Deposit Date:', emd_data.get('deposit_date', 'N/A')],
            ['Refund Date:', emd_data.get('refund_date', 'N/A')],
            ['Interest Rate:', f"{emd_data.get('interest_rate', 0):.2f}% per annum"],
            ['Interest Amount:', f"₹{emd_data.get('interest_amount', 0):,.2f}"],
            ['Total Refund Amount:', f"₹{emd_data.get('total_refund', 0):,.2f}"]
        ]
        
        details_table = Table(emd_details, colWidths=[2.5*inch, 3.5*inch])
        details_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F0F2F6')),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#1A8A16')),
            ('TEXTCOLOR', (0, -1), (-1, -1), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -2), 'Helvetica'),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(details_table)
        story.append(Spacer(1, 20))
        
        # Calculation details
        if emd_data.get('days_held'):
            calc_para = Paragraph(
                f"<b>Calculation Details:</b><br/>"
                f"EMD Amount: ₹{emd_data.get('emd_amount', 0):,.2f}<br/>"
                f"Days Held: {emd_data.get('days_held', 0)} days<br/>"
                f"Interest Calculation: ₹{emd_data.get('emd_amount', 0):,.2f} × {emd_data.get('interest_rate', 0):.2f}% × {emd_data.get('days_held', 0)}/365<br/>"
                f"Interest Amount: ₹{emd_data.get('interest_amount', 0):,.2f}<br/>"
                f"Total Refund: ₹{emd_data.get('total_refund', 0):,.2f}",
                self.styles['PWDNormal']
            )
            story.append(calc_para)
            story.append(Spacer(1, 20))
        
        # Certification
        certification = Paragraph(
            "This is to certify that the above EMD refund has been calculated correctly and is approved for payment.",
            self.styles['PWDNormal']
        )
        story.append(certification)
        story.append(Spacer(1, 30))
        
        # Signature section
        signature_data = [
            ['Accounts Officer', 'Assistant Engineer', 'Executive Engineer'],
            ['_' * 15, '_' * 15, '_' * 15],
            ['Date: _________', 'Date: _________', 'Date: _________']
        ]
        
        signature_table = Table(signature_data, colWidths=[2.2*inch, 2.2*inch, 2.2*inch])
        signature_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6)
        ]))
        
        story.append(signature_table)
        
        # Footer
        footer = Paragraph(
            f"Generated on: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br/>PWD Tools - EMD Refund Calculator",
            self.styles['Normal']
        )
        story.append(Spacer(1, 30))
        story.append(footer)
        
        # Build PDF
        doc.build(story)
        
        if output_path is None:
            buffer.seek(0)
            return buffer.getvalue()
        
        return output_path
    
    def create_project_report_pdf(self, project_data, output_path=None):
        """Generate PDF for project financial report"""
        if output_path is None:
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
        else:
            doc = SimpleDocTemplate(output_path, pagesize=A4)
        
        story = []
        
        # Header
        header = Paragraph("PROJECT FINANCIAL REPORT<br/>PWD - Public Works Department", self.styles['PWDHeader'])
        story.append(header)
        story.append(Spacer(1, 20))
        
        # Project details
        project_details = [
            ['Project Name:', project_data.get('project_name', 'N/A')],
            ['Project Code:', project_data.get('project_code', 'N/A')],
            ['Contractor Name:', project_data.get('contractor_name', 'N/A')],
            ['Agreement Amount:', f"₹{project_data.get('agreement_amount', 0):,.2f}"],
            ['Work Done Amount:', f"₹{project_data.get('work_done_amount', 0):,.2f}"],
            ['Payments Made:', f"₹{project_data.get('payments_made', 0):,.2f}"],
            ['Outstanding Amount:', f"₹{project_data.get('outstanding_amount', 0):,.2f}"]
        ]
        
        details_table = Table(project_details, colWidths=[2.5*inch, 3.5*inch])
        details_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F0F2F6')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(details_table)
        story.append(Spacer(1, 20))
        
        # Progress summary
        if 'physical_progress' in project_data:
            progress_header = Paragraph("Progress Summary", self.styles['PWDSubHeader'])
            story.append(progress_header)
            
            progress_data = [
                ['Physical Progress:', f"{project_data.get('physical_progress', 0):.1f}%"],
                ['Financial Progress:', f"{project_data.get('financial_progress', 0):.1f}%"],
                ['Time Progress:', f"{project_data.get('time_progress', 0):.1f}%"]
            ]
            
            progress_table = Table(progress_data, colWidths=[2.5*inch, 3.5*inch])
            progress_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F0F2F6')),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            story.append(progress_table)
            story.append(Spacer(1, 20))
        
        # Footer
        footer = Paragraph(
            f"Report Generated on: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br/>"
            f"PWD Tools - Financial Analysis Module",
            self.styles['Normal']
        )
        story.append(footer)
        
        # Build PDF
        doc.build(story)
        
        if output_path is None:
            buffer.seek(0)
            return buffer.getvalue()
        
        return output_path

# Fallback class if ReportLab is not available
class PDFGeneratorFallback:
    """Fallback PDF generator when ReportLab is not available"""
    
    def __init__(self):
        self.available = False
    
    def create_bill_pdf(self, bill_data, output_path=None):
        raise NotImplementedError("PDF generation requires ReportLab. Install with: pip install reportlab")
    
    def create_emd_refund_pdf(self, emd_data, output_path=None):
        raise NotImplementedError("PDF generation requires ReportLab. Install with: pip install reportlab")
    
    def create_project_report_pdf(self, project_data, output_path=None):
        raise NotImplementedError("PDF generation requires ReportLab. Install with: pip install reportlab")

# Factory function
def get_pdf_generator():
    """Get PDF generator instance"""
    if REPORTLAB_AVAILABLE:
        return PDFGenerator()
    else:
        return PDFGeneratorFallback()

# Utility functions for text-based reports (always available)
def generate_text_bill_report(bill_data):
    """Generate text-based bill report"""
    report = f"""
PWD BILL REPORT
===============

Bill Number: {bill_data.get('bill_number', 'N/A')}
Bill Date: {bill_data.get('bill_date', 'N/A')}
Contractor: {bill_data.get('contractor_name', 'N/A')}
Project: {bill_data.get('project_name', 'N/A')}

FINANCIAL DETAILS:
Bill Amount: ₹{bill_data.get('bill_amount', 0):,.2f}
Total Deductions: ₹{sum([d.get('amount', 0) for d in bill_data.get('deductions', [])]):,.2f}
Net Payable: ₹{bill_data.get('bill_amount', 0) - sum([d.get('amount', 0) for d in bill_data.get('deductions', [])]):,.2f}

Generated on: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
PWD Tools - Bill Generator
"""
    return report

def generate_text_emd_report(emd_data):
    """Generate text-based EMD refund report"""
    report = f"""
EMD REFUND CERTIFICATE
=====================

Tender Number: {emd_data.get('tender_number', 'N/A')}
Contractor: {emd_data.get('contractor_name', 'N/A')}
EMD Amount: ₹{emd_data.get('emd_amount', 0):,.2f}
Interest Rate: {emd_data.get('interest_rate', 0):.2f}% per annum
Total Refund: ₹{emd_data.get('total_refund', 0):,.2f}

Generated on: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
PWD Tools - EMD Refund Calculator
"""
    return report
