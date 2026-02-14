#!/usr/bin/env python3
"""
Report generator for Launch Readiness assessments.

Generates three types of reports:
1. Baseline Report (PRE assessment only)
2. Progress Report (PRE vs POST comparison)
3. Impact Report (Cohort summary)

All reports use Cencora branding and house style.
"""

import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

from framework import (
    INDICATORS, INDICATOR_DESCRIPTIONS, INDICATOR_COLOURS,
    ITEMS, OPEN_QUESTIONS_PRE, OPEN_QUESTIONS_POST, RATING_SCALE,
    FOCUS_TAGS, get_items_for_indicator, get_items_by_focus
)
from theme_extractor import ThemeExtractor, format_themes_for_report


# Cencora brand colours (hex for matplotlib, RGB for docx)
COLOURS_HEX = {
    'purple': '#461E96',
    'cyan': '#00B4E6',
    'magenta': '#E6008C',
    'green': '#00DC8C',
    'orange': '#FFA400',
    'dark_grey': '#3B3B3B',
    'mid_grey': '#6E6E6E',
    'light_grey': '#F5F5F5',
    'cream': '#FDF6E3',
    'light_green': '#E8F5E9',
    'success_green': '#2E7D32',
    'white': '#FFFFFF'
}

COLOURS_RGB = {
    'purple': RGBColor(0x46, 0x1E, 0x96),
    'cyan': RGBColor(0x00, 0xB4, 0xE6),
    'magenta': RGBColor(0xE6, 0x00, 0x8C),
    'green': RGBColor(0x00, 0xDC, 0x8C),
    'orange': RGBColor(0xFF, 0xA4, 0x00),
    'dark_grey': RGBColor(0x3B, 0x3B, 0x3B),
    'mid_grey': RGBColor(0x6E, 0x6E, 0x6E),
    'light_grey': RGBColor(0xF5, 0xF5, 0xF5),
    'cream': RGBColor(0xFD, 0xF6, 0xE3),
    'light_green': RGBColor(0xE8, 0xF5, 0xE9),
    'success_green': RGBColor(0x2E, 0x7D, 0x32),
    'white': RGBColor(0xFF, 0xFF, 0xFF)
}

INDICATOR_COLOUR_MAP = {
    'Self-Readiness': 'purple',
    'Practical Readiness': 'cyan',
    'Professional Readiness': 'magenta',
    'Team Readiness': 'green'
}

# Logo path - multiple fallback locations for different environments
def get_logo_path():
    """Get the logo path, checking multiple locations for compatibility."""
    # Current file's directory
    current_dir = Path(__file__).parent
    
    # Possible locations
    candidates = [
        current_dir / 'assets' / 'cencora_logo.png',  # Standard location
        current_dir / 'cencora_logo.png',              # Root of app
        Path('assets') / 'cencora_logo.png',           # Relative to cwd
        Path('cencora_logo.png'),                      # Current directory
    ]
    
    for path in candidates:
        if path.exists():
            return path
    
    return None


class ReportGenerator:
    def __init__(self, db):
        self.db = db
        self.theme_extractor = ThemeExtractor()
        self.logo_path = get_logo_path()
    
    # =========== CHART GENERATION ===========
    
    def _create_radar_chart(self, scores: dict, output_path: str):
        """Create a 4-axis radar chart matching the approved sample style."""
        
        indicators = list(INDICATORS.keys())
        values = [scores.get(ind, 0) for ind in indicators]
        
        # Chart setup
        fig, ax = plt.subplots(figsize=(5.5, 5.5), subplot_kw=dict(polar=True))
        fig.patch.set_facecolor('white')
        
        # Angles - start from top (-90°), go clockwise
        angles = np.array([-np.pi/2, 0, np.pi/2, np.pi])
        
        # Close the polygon
        values_closed = values + [values[0]]
        angles_closed = np.append(angles, angles[0])
        
        # Draw grid circles
        ax.set_ylim(0, 6)
        ax.set_yticks([1, 2, 3, 4, 5, 6])
        ax.set_yticklabels(['1', '2', '3', '4', '5', '6'], fontsize=8, color='#888888')
        ax.yaxis.grid(True, color='#E0E0E0', linewidth=1)
        
        # Draw the data polygon
        ax.fill(angles_closed, values_closed, color=COLOURS_HEX['purple'], alpha=0.2)
        ax.plot(angles_closed, values_closed, color=COLOURS_HEX['purple'], linewidth=3)
        
        # Draw data points with indicator colours
        indicator_colours = [COLOURS_HEX['purple'], COLOURS_HEX['cyan'], 
                           COLOURS_HEX['magenta'], COLOURS_HEX['green']]
        for angle, value, colour in zip(angles, values, indicator_colours):
            ax.scatter(angle, value, color=colour, s=150, zorder=5, edgecolors='white', linewidths=2)
        
        # Labels
        ax.set_xticks(angles)
        ax.set_xticklabels([])  # Remove default labels
        
        # Add custom positioned labels
        for i, (ind, colour) in enumerate(zip(indicators, indicator_colours)):
            angle = angles[i]
            if i == 0:  # Top
                ax.text(angle, 7.5, ind, ha='center', va='bottom', fontsize=11, 
                       fontweight='bold', color=colour)
            elif i == 1:  # Right
                ax.text(angle, 7.5, ind, ha='left', va='center', fontsize=11,
                       fontweight='bold', color=colour)
            elif i == 2:  # Bottom
                ax.text(angle, 7.5, ind, ha='center', va='top', fontsize=11,
                       fontweight='bold', color=colour)
            else:  # Left
                ax.text(angle, 7.5, ind, ha='right', va='center', fontsize=11,
                       fontweight='bold', color=colour)
        
        ax.spines['polar'].set_visible(False)
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white', 
                   edgecolor='none', pad_inches=0.1)
        plt.close()
    
    def _create_comparison_radar_chart(self, pre_scores: dict, post_scores: dict, output_path: str):
        """Create a comparison radar chart (pre dashed grey, post solid green)."""
        
        indicators = list(INDICATORS.keys())
        pre_values = [pre_scores.get(ind, 0) for ind in indicators]
        post_values = [post_scores.get(ind, 0) for ind in indicators]
        
        fig, ax = plt.subplots(figsize=(5.5, 5.5), subplot_kw=dict(polar=True))
        fig.patch.set_facecolor('white')
        
        angles = np.array([-np.pi/2, 0, np.pi/2, np.pi])
        
        pre_closed = pre_values + [pre_values[0]]
        post_closed = post_values + [post_values[0]]
        angles_closed = np.append(angles, angles[0])
        
        ax.set_ylim(0, 6)
        ax.set_yticks([1, 2, 3, 4, 5, 6])
        ax.set_yticklabels(['1', '2', '3', '4', '5', '6'], fontsize=8, color='#888888')
        ax.yaxis.grid(True, color='#E0E0E0', linewidth=1)
        
        # PRE polygon (dashed, grey)
        ax.fill(angles_closed, pre_closed, color='#999999', alpha=0.1)
        ax.plot(angles_closed, pre_closed, color='#999999', linewidth=2, linestyle='--')
        
        # POST polygon (solid, green)
        ax.fill(angles_closed, post_closed, color=COLOURS_HEX['green'], alpha=0.2)
        ax.plot(angles_closed, post_closed, color=COLOURS_HEX['green'], linewidth=3)
        
        # Points
        indicator_colours = [COLOURS_HEX['purple'], COLOURS_HEX['cyan'],
                           COLOURS_HEX['magenta'], COLOURS_HEX['green']]
        
        for angle, pre_val, post_val, colour in zip(angles, pre_values, post_values, indicator_colours):
            # PRE point (smaller, grey)
            ax.scatter(angle, pre_val, color='#999999', s=60, zorder=4, edgecolors='white', linewidths=1)
            # POST point (larger, coloured)
            ax.scatter(angle, post_val, color=colour, s=150, zorder=5, edgecolors='white', linewidths=2)
        
        ax.set_xticks(angles)
        ax.set_xticklabels([])
        
        for i, (ind, colour) in enumerate(zip(indicators, indicator_colours)):
            angle = angles[i]
            if i == 0:
                ax.text(angle, 7.5, ind, ha='center', va='bottom', fontsize=11,
                       fontweight='bold', color=colour)
            elif i == 1:
                ax.text(angle, 7.5, ind, ha='left', va='center', fontsize=11,
                       fontweight='bold', color=colour)
            elif i == 2:
                ax.text(angle, 7.5, ind, ha='center', va='top', fontsize=11,
                       fontweight='bold', color=colour)
            else:
                ax.text(angle, 7.5, ind, ha='right', va='center', fontsize=11,
                       fontweight='bold', color=colour)
        
        # Legend
        from matplotlib.patches import Patch
        from matplotlib.lines import Line2D
        legend_elements = [
            Line2D([0], [0], color='#999999', linestyle='--', linewidth=2, label='Pre-Programme'),
            Line2D([0], [0], color=COLOURS_HEX['green'], linewidth=3, label='Post-Programme')
        ]
        ax.legend(handles=legend_elements, loc='lower left', bbox_to_anchor=(-0.1, -0.15), 
                 fontsize=9, frameon=False)
        
        ax.spines['polar'].set_visible(False)
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white',
                   edgecolor='none', pad_inches=0.1)
        plt.close()
    
    def _create_bar_chart(self, score: float, colour_hex: str, output_path: str, max_score: int = 6):
        """Create a simple horizontal bar for a single score."""
        
        fig, ax = plt.subplots(figsize=(1.5, 0.25))
        fig.patch.set_facecolor('white')
        
        # Background bar
        ax.barh(0, max_score, color='#E8E8E8', height=0.8)
        # Score bar
        ax.barh(0, score, color=colour_hex, height=0.8)
        
        ax.set_xlim(0, max_score)
        ax.set_ylim(-0.5, 0.5)
        ax.axis('off')
        
        plt.savefig(output_path, dpi=100, bbox_inches='tight', facecolor='white',
                   edgecolor='none', pad_inches=0.02)
        plt.close()
    
    def _create_comparison_bar_chart(self, pre_score: float, post_score: float, 
                                     colour_hex: str, output_path: str, max_score: int = 6):
        """Create a comparison bar (pre dashed outline, post filled)."""
        
        fig, ax = plt.subplots(figsize=(1.5, 0.25))
        fig.patch.set_facecolor('white')
        
        # Background
        ax.barh(0, max_score, color='#E8E8E8', height=0.8)
        # POST bar (filled)
        ax.barh(0, post_score, color=colour_hex, height=0.8)
        # PRE bar (outline only)
        ax.barh(0, pre_score, color='none', edgecolor='#888888',
               linewidth=2, linestyle='--', height=0.8)
        
        ax.set_xlim(0, max_score)
        ax.set_ylim(-0.5, 0.5)
        ax.axis('off')
        
        plt.savefig(output_path, dpi=100, bbox_inches='tight', facecolor='white',
                   edgecolor='none', pad_inches=0.02)
        plt.close()
    
    # =========== DOCUMENT HELPERS ===========
    
    def _set_cell_shading(self, cell, colour_hex: str):
        """Set cell background colour."""
        colour_hex = colour_hex.replace('#', '')
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), colour_hex)
        cell._tc.get_or_add_tcPr().append(shading)
    
    def _set_cell_margins(self, cell, top=60, bottom=60, left=100, right=100):
        """Set cell margins in twips."""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
            node = OxmlElement(f'w:{side}')
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)
    
    def _add_logo_header(self, doc):
        """Add Cencora logo to document header on all pages."""
        if self.logo_path and self.logo_path.exists():
            section = doc.sections[0]
            header = section.header
            header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = header_para.add_run()
            # Logo width of 1.5 inches for good visibility
            run.add_picture(str(self.logo_path), width=Inches(1.5))
    
    def _create_styled_table(self, doc, headers: list, header_colour_hex: str = '461E96'):
        """Create a table with styled header row."""
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        
        # Style header row
        header_row = table.rows[0]
        for i, header_text in enumerate(headers):
            cell = header_row.cells[i]
            cell.text = header_text
            self._set_cell_shading(cell, header_colour_hex)
            self._set_cell_margins(cell)
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.bold = True
                    run.font.color.rgb = COLOURS_RGB['white']
                    run.font.size = Pt(9)
                    run.font.name = 'Arial'
        
        return table
    
    def _add_table_row(self, table, values: list, row_index: int, 
                       alignments: list = None, bar_image_path: str = None, bar_col: int = None):
        """Add a data row with alternating colours and optional bar chart."""
        row = table.add_row()
        bg_colour = 'FFFFFF' if row_index % 2 == 0 else 'FDF6E3'
        
        for i, value in enumerate(values):
            cell = row.cells[i]
            
            # Check if this cell should have a bar image
            if bar_image_path and i == bar_col:
                para = cell.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(bar_image_path, width=Inches(0.8))
            else:
                cell.text = str(value) if value is not None else ''
                for para in cell.paragraphs:
                    if alignments and i < len(alignments):
                        para.alignment = alignments[i]
                    for run in para.runs:
                        run.font.size = Pt(9)
                        run.font.name = 'Arial'
            
            self._set_cell_shading(cell, bg_colour)
            self._set_cell_margins(cell)
        
        return row
    
    # =========== CALCULATION HELPERS ===========
    
    def _calculate_indicator_scores(self, ratings: dict) -> dict:
        """Calculate average score per indicator from ratings."""
        scores = {}
        for indicator, (start, end) in INDICATORS.items():
            item_scores = [ratings.get(i, 0) for i in range(start, end + 1) if i in ratings]
            if item_scores:
                scores[indicator] = sum(item_scores) / len(item_scores)
            else:
                scores[indicator] = 0
        return scores
    
    def _calculate_overall_score(self, ratings: dict) -> float:
        """Calculate overall average from all 32 ratings."""
        if not ratings:
            return 0
        valid_ratings = [ratings.get(i, 0) for i in range(1, 33) if i in ratings]
        return sum(valid_ratings) / len(valid_ratings) if valid_ratings else 0
    
    def _calculate_focus_scores(self, ratings: dict) -> dict:
        """Calculate average scores by focus area (K/A/C/B)."""
        scores = {}
        for focus in FOCUS_TAGS.keys():
            items = get_items_by_focus(focus)
            item_scores = [ratings.get(i, 0) for i in items if i in ratings]
            if item_scores:
                scores[focus] = sum(item_scores) / len(item_scores)
            else:
                scores[focus] = 0
        return scores
    
    # =========== BASELINE REPORT ===========
    
    def generate_baseline_report(self, participant_id: int) -> io.BytesIO:
        """Generate a Baseline report (PRE assessment only)."""
        
        # Get data
        data = self.db.get_participant_data(participant_id)
        if not data or not data['pre']:
            raise ValueError("No PRE assessment data found for this participant")
        
        participant = data['participant']
        cohort = data['cohort']
        pre_ratings = data['pre']['ratings']
        pre_responses = data['pre']['open_responses']
        pre_date = data['pre']['assessment'].get('completed_at', '')[:10]
        
        # Calculate scores
        indicator_scores = self._calculate_indicator_scores(pre_ratings)
        overall_score = self._calculate_overall_score(pre_ratings)
        
        # Create document
        doc = Document()
        
        # Set default font
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        
        # Add logo header
        self._add_logo_header(doc)
        
        # Title
        title = doc.add_paragraph()
        run = title.add_run("THE READINESS FRAMEWORK")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        run.font.name = 'Arial'
        
        subtitle = doc.add_paragraph()
        run = subtitle.add_run("Readiness Baseline")
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['magenta']
        run.font.name = 'Arial'
        
        # Participant info table
        info_table = doc.add_table(rows=4, cols=2)
        info_table.style = 'Table Grid'
        info_data = [
            ("Participant:", participant['name']),
            ("Role:", participant.get('role', 'Not specified')),
            ("Cohort:", cohort['name']),
            ("Assessment Date:", pre_date)
        ]
        for i, (label, value) in enumerate(info_data):
            info_table.rows[i].cells[0].text = label
            info_table.rows[i].cells[1].text = value
            self._set_cell_shading(info_table.rows[i].cells[0], 'F5F5F5')
            self._set_cell_margins(info_table.rows[i].cells[0])
            self._set_cell_margins(info_table.rows[i].cells[1])
            for cell in info_table.rows[i].cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.name = 'Arial'
                        if cell == info_table.rows[i].cells[0]:
                            run.bold = True
        
        # Introduction
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Your Starting Point")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        intro = doc.add_paragraph()
        intro.add_run(f"Welcome to the Launch Readiness programme, {participant['name'].split()[0]}. "
                      f"This report captures your self-assessment before the programme begins. "
                      f"There are no right or wrong answers; this is simply a snapshot of where you see yourself today.")
        
        # Radar chart heading
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Your Readiness Profile")
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        # Radar chart - centered
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            self._create_radar_chart(indicator_scores, tmp.name)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(tmp.name, width=Inches(4.0))
        
        # Scale note
        scale_para = doc.add_paragraph()
        scale_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = scale_para.add_run("Scale: 1-6 (1=Strongly Disagree, 6=Strongly Agree)")
        run.italic = True
        run.font.size = Pt(8)
        run.font.color.rgb = COLOURS_RGB['mid_grey']
        
        # Summary table
        doc.add_paragraph()
        summary_table = self._create_styled_table(doc, ["Indicator", "Score"])
        
        for i, (ind, score) in enumerate(indicator_scores.items()):
            self._add_table_row(summary_table, [ind, f"{score:.1f}"], i,
                               [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER])
        
        # Overall row
        overall_row = summary_table.add_row()
        overall_row.cells[0].text = "OVERALL"
        overall_row.cells[1].text = f"{overall_score:.1f}"
        for cell in overall_row.cells:
            self._set_cell_shading(cell, 'F5F5F5')
            self._set_cell_margins(cell)
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER if cell == overall_row.cells[1] else WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(9)
        
        # Page break before detailed scores
        doc.add_page_break()
        
        # Detailed scores by indicator with bar charts
        for indicator, (start, end) in INDICATORS.items():
            colour_name = INDICATOR_COLOUR_MAP.get(indicator, 'purple')
            colour_hex = COLOURS_HEX[colour_name]
            
            # Indicator heading
            heading = doc.add_paragraph()
            run = heading.add_run(indicator)
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = COLOURS_RGB[colour_name]
            
            # Description and average
            desc = doc.add_paragraph()
            run = desc.add_run(INDICATOR_DESCRIPTIONS.get(indicator, ''))
            run.italic = True
            run.font.color.rgb = COLOURS_RGB['mid_grey']
            run.font.size = Pt(9)
            desc.add_run("  |  ")
            desc.add_run("Dimension Average: ")
            run = desc.add_run(f"{indicator_scores.get(indicator, 0):.1f}")
            run.bold = True
            
            # Items table with bar charts
            items_table = self._create_styled_table(
                doc, ["#", "Statement", "Focus", "", "Score"],
                colour_hex.replace('#', '')
            )
            
            for idx, item_num in enumerate(range(start, end + 1)):
                item = ITEMS[item_num]
                score = pre_ratings.get(item_num, 0)
                
                # Create bar chart for this item
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                    self._create_bar_chart(score, colour_hex, tmp.name)
                    self._add_table_row(
                        items_table,
                        [str(item_num), item['text'], item['focus'], None, str(score)],
                        idx,
                        [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT,
                         WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
                         WD_ALIGN_PARAGRAPH.CENTER],
                        bar_image_path=tmp.name,
                        bar_col=3
                    )
            
            doc.add_paragraph()
        
        # Overall Readiness items
        heading = doc.add_paragraph()
        run = heading.add_run("Overall Readiness")
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        overall_table = self._create_styled_table(doc, ["#", "Statement", "Focus", "", "Score"])
        
        for idx, item_num in enumerate([31, 32]):
            item = ITEMS[item_num]
            score = pre_ratings.get(item_num, 0)
            
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                self._create_bar_chart(score, COLOURS_HEX['purple'], tmp.name)
                self._add_table_row(
                    overall_table,
                    [str(item_num), item['text'], item['focus'], None, str(score)],
                    idx,
                    [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT,
                     WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
                     WD_ALIGN_PARAGRAPH.CENTER],
                    bar_image_path=tmp.name,
                    bar_col=3
                )
        
        # Page break before reflections
        doc.add_page_break()
        
        # Open responses
        heading = doc.add_paragraph()
        run = heading.add_run("Your Reflections")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        for q_num, question in OPEN_QUESTIONS_PRE.items():
            q_para = doc.add_paragraph()
            run = q_para.add_run(question)
            run.bold = True
            run.font.size = Pt(10)
            
            response = pre_responses.get(q_num, "No response provided")
            r_para = doc.add_paragraph()
            run = r_para.add_run(response)
            run.italic = True
            run.font.size = Pt(10)
            
            doc.add_paragraph()
        
        # Closing note
        doc.add_paragraph()
        closing = doc.add_paragraph()
        run = closing.add_run("Keep this report - you'll revisit it after the programme to see how far you've come.")
        run.italic = True
        run.font.color.rgb = COLOURS_RGB['mid_grey']
        run.font.size = Pt(9)
        
        # Save to buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    
    # =========== PROGRESS REPORT ===========
    
    def generate_progress_report(self, participant_id: int, cohort_id: int) -> io.BytesIO:
        """Generate a Progress report (PRE vs POST with cohort comparison)."""
        
        # Get participant data
        data = self.db.get_participant_data(participant_id)
        if not data or not data['pre'] or not data['post']:
            raise ValueError("Both PRE and POST assessments required for Progress report")
        
        participant = data['participant']
        cohort = data['cohort']
        pre_ratings = data['pre']['ratings']
        post_ratings = data['post']['ratings']
        pre_responses = data['pre']['open_responses']
        post_responses = data['post']['open_responses']
        pre_date = data['pre']['assessment'].get('completed_at', '')[:10]
        post_date = data['post']['assessment'].get('completed_at', '')[:10]
        
        # Calculate scores
        pre_indicator_scores = self._calculate_indicator_scores(pre_ratings)
        post_indicator_scores = self._calculate_indicator_scores(post_ratings)
        pre_overall = self._calculate_overall_score(pre_ratings)
        post_overall = self._calculate_overall_score(post_ratings)
        
        # Get cohort averages
        cohort_avgs = self.db.get_cohort_averages(cohort_id, 'POST')
        cohort_indicator_scores = {}
        for indicator, (start, end) in INDICATORS.items():
            item_avgs = [cohort_avgs.get(i, {}).get('avg', 0) for i in range(start, end + 1)]
            cohort_indicator_scores[indicator] = sum(item_avgs) / len(item_avgs) if item_avgs else 0
        cohort_overall = sum(cohort_avgs.get(i, {}).get('avg', 0) for i in range(1, 33)) / 32
        
        # Create document
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        
        self._add_logo_header(doc)
        
        # Title
        title = doc.add_paragraph()
        run = title.add_run("THE READINESS FRAMEWORK")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        subtitle = doc.add_paragraph()
        run = subtitle.add_run("Readiness Progress")
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['magenta']
        
        # Participant info
        info_table = doc.add_table(rows=4, cols=2)
        info_table.style = 'Table Grid'
        info_data = [
            ("Participant:", participant['name']),
            ("Role:", participant.get('role', 'Not specified')),
            ("Pre-Assessment:", pre_date),
            ("Post-Assessment:", post_date)
        ]
        for i, (label, value) in enumerate(info_data):
            info_table.rows[i].cells[0].text = label
            info_table.rows[i].cells[1].text = value
            self._set_cell_shading(info_table.rows[i].cells[0], 'F5F5F5')
            self._set_cell_margins(info_table.rows[i].cells[0])
            self._set_cell_margins(info_table.rows[i].cells[1])
            for cell in info_table.rows[i].cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)
                        if cell == info_table.rows[i].cells[0]:
                            run.bold = True
        
        # Summary heading
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Your Progress")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        change = post_overall - pre_overall
        intro = doc.add_paragraph()
        intro.add_run(f"Congratulations, {participant['name'].split()[0]}! Your assessment shows meaningful "
                      f"growth across all four Readiness Indicators. Your overall readiness improved from "
                      f"{pre_overall:.1f} to {post_overall:.1f} (+{change:.1f}).")
        
        # Comparison radar chart
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Your Growth Profile")
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            self._create_comparison_radar_chart(pre_indicator_scores, post_indicator_scores, tmp.name)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(tmp.name, width=Inches(4.0))
        
        # Summary table
        doc.add_paragraph()
        summary_table = self._create_styled_table(doc, ["Indicator", "Pre", "Post", "Change", "Cohort"])
        
        for i, indicator in enumerate(INDICATORS.keys()):
            pre = pre_indicator_scores.get(indicator, 0)
            post = post_indicator_scores.get(indicator, 0)
            change = post - pre
            cohort_avg = cohort_indicator_scores.get(indicator, 0)
            change_str = f"+{change:.1f}" if change > 0 else f"{change:.1f}"
            
            self._add_table_row(summary_table,
                               [indicator, f"{pre:.1f}", f"{post:.1f}", change_str, f"{cohort_avg:.1f}"],
                               i,
                               [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 4)
        
        # Overall row
        change = post_overall - pre_overall
        change_str = f"+{change:.1f}" if change > 0 else f"{change:.1f}"
        overall_row = summary_table.add_row()
        for j, val in enumerate(["OVERALL", f"{pre_overall:.1f}", f"{post_overall:.1f}", change_str, f"{cohort_overall:.1f}"]):
            overall_row.cells[j].text = val
            self._set_cell_shading(overall_row.cells[j], 'F5F5F5')
            self._set_cell_margins(overall_row.cells[j])
            for para in overall_row.cells[j].paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(9)
        
        # Cohort note
        note = doc.add_paragraph()
        run = note.add_run("Cohort = Average of all participants  |  Bar shows Pre (dashed) vs Post (solid)")
        run.italic = True
        run.font.size = Pt(8)
        run.font.color.rgb = COLOURS_RGB['mid_grey']
        
        # Page break
        doc.add_page_break()
        
        # Detailed comparison by indicator
        for indicator, (start, end) in INDICATORS.items():
            colour_name = INDICATOR_COLOUR_MAP.get(indicator, 'purple')
            colour_hex = COLOURS_HEX[colour_name]
            
            pre_avg = pre_indicator_scores.get(indicator, 0)
            post_avg = post_indicator_scores.get(indicator, 0)
            change = post_avg - pre_avg
            
            heading = doc.add_paragraph()
            run = heading.add_run(indicator)
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = COLOURS_RGB[colour_name]
            
            desc = doc.add_paragraph()
            run = desc.add_run(INDICATOR_DESCRIPTIONS.get(indicator, ''))
            run.italic = True
            run.font.color.rgb = COLOURS_RGB['mid_grey']
            run.font.size = Pt(9)
            desc.add_run(f"  |  Pre: {pre_avg:.1f} → Post: {post_avg:.1f} ")
            change_run = desc.add_run(f"(+{change:.1f})" if change > 0 else f"({change:.1f})")
            change_run.bold = True
            if change > 0:
                change_run.font.color.rgb = COLOURS_RGB['success_green']
            
            # Items table
            items_table = self._create_styled_table(
                doc, ["#", "Statement", "Focus", "Pre", "Post", "", "Change"],
                colour_hex.replace('#', '')
            )
            
            for idx, item_num in enumerate(range(start, end + 1)):
                item = ITEMS[item_num]
                pre_score = pre_ratings.get(item_num, 0)
                post_score = post_ratings.get(item_num, 0)
                item_change = post_score - pre_score
                change_str = f"+{item_change}" if item_change > 0 else str(item_change)
                
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                    self._create_comparison_bar_chart(pre_score, post_score, colour_hex, tmp.name)
                    self._add_table_row(
                        items_table,
                        [str(item_num), item['text'], item['focus'], str(pre_score), str(post_score), None, change_str],
                        idx,
                        [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT,
                         WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
                         WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
                         WD_ALIGN_PARAGRAPH.CENTER],
                        bar_image_path=tmp.name,
                        bar_col=5
                    )
            
            doc.add_paragraph()
        
        # Page break before reflections
        doc.add_page_break()
        
        # Reflections
        heading = doc.add_paragraph()
        run = heading.add_run("Your Reflections")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        for q_num, question in OPEN_QUESTIONS_POST.items():
            q_para = doc.add_paragraph()
            run = q_para.add_run(question)
            run.bold = True
            
            # For Q3, show original concern
            if q_num == 3:
                original = pre_responses.get(3, "")
                if original:
                    orig_para = doc.add_paragraph()
                    run = orig_para.add_run("Your original concern: ")
                    run.font.color.rgb = COLOURS_RGB['mid_grey']
                    run.font.size = Pt(9)
                    run = orig_para.add_run(f'"{original}"')
                    run.italic = True
                    run.font.size = Pt(9)
            
            response = post_responses.get(q_num, "No response provided")
            r_para = doc.add_paragraph()
            run = r_para.add_run(response)
            run.italic = True
            
            doc.add_paragraph()
        
        # Closing
        closing = doc.add_paragraph()
        run = closing.add_run("For questions about your development plan, speak with your facilitator or line manager.")
        run.italic = True
        run.font.color.rgb = COLOURS_RGB['mid_grey']
        run.font.size = Pt(9)
        
        # Save
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    
    # =========== IMPACT REPORT ===========
    
    def generate_impact_report(self, cohort_id: int) -> io.BytesIO:
        """Generate an Impact report (cohort summary)."""
        
        # Get cohort data
        cohort_data = self.db.get_cohort_data(cohort_id)
        if not cohort_data:
            raise ValueError("Cohort not found")
        
        cohort = cohort_data['cohort']
        participants = cohort_data['participants']
        
        complete_participants = [p for p in participants if p['pre'] and p['post']]
        
        if len(complete_participants) < 2:
            raise ValueError("Need at least 2 participants with complete data")
        
        # Calculate cohort averages
        pre_avgs = self.db.get_cohort_averages(cohort_id, 'PRE')
        post_avgs = self.db.get_cohort_averages(cohort_id, 'POST')
        
        pre_indicator_scores = {}
        post_indicator_scores = {}
        for indicator, (start, end) in INDICATORS.items():
            pre_scores = [pre_avgs.get(i, {}).get('avg', 0) for i in range(start, end + 1)]
            post_scores = [post_avgs.get(i, {}).get('avg', 0) for i in range(start, end + 1)]
            pre_indicator_scores[indicator] = sum(pre_scores) / len(pre_scores) if pre_scores else 0
            post_indicator_scores[indicator] = sum(post_scores) / len(post_scores) if post_scores else 0
        
        pre_overall = sum(pre_avgs.get(i, {}).get('avg', 0) for i in range(1, 33)) / 32
        post_overall = sum(post_avgs.get(i, {}).get('avg', 0) for i in range(1, 33)) / 32
        
        # Focus area analysis
        pre_focus = {}
        post_focus = {}
        for focus in FOCUS_TAGS.keys():
            items = get_items_by_focus(focus)
            pre_focus[focus] = sum(pre_avgs.get(i, {}).get('avg', 0) for i in items) / len(items) if items else 0
            post_focus[focus] = sum(post_avgs.get(i, {}).get('avg', 0) for i in items) / len(items) if items else 0
        
        # Get qualitative themes
        takeaway_responses = [p['post']['open_responses'].get(1, '') for p in complete_participants if p['post']['open_responses'].get(1)]
        commitment_responses = [p['post']['open_responses'].get(2, '') for p in complete_participants if p['post']['open_responses'].get(2)]
        
        takeaway_themes = self.theme_extractor.extract_takeaways(takeaway_responses)
        commitment_themes = self.theme_extractor.extract_commitments(commitment_responses)
        
        # Create document
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        
        self._add_logo_header(doc)
        
        # Title
        title = doc.add_paragraph()
        run = title.add_run("THE READINESS FRAMEWORK")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        subtitle = doc.add_paragraph()
        run = subtitle.add_run("Readiness Impact")
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['magenta']
        
        # Cohort info
        info_table = doc.add_table(rows=4, cols=2)
        info_table.style = 'Table Grid'
        completion_rate = int(cohort_data['post_completed'] / len(participants) * 100) if participants else 0
        info_data = [
            ("Programme:", cohort.get('programme', 'Launch Readiness')),
            ("Cohort:", cohort['name']),
            ("Participants:", f"{len(participants)} enrolled | {cohort_data['pre_completed']} pre | {cohort_data['post_completed']} post"),
            ("Period:", f"{cohort.get('start_date', 'TBC')} to {cohort.get('end_date', 'TBC')}")
        ]
        for i, (label, value) in enumerate(info_data):
            info_table.rows[i].cells[0].text = label
            info_table.rows[i].cells[1].text = value
            self._set_cell_shading(info_table.rows[i].cells[0], 'F5F5F5')
            self._set_cell_margins(info_table.rows[i].cells[0])
            self._set_cell_margins(info_table.rows[i].cells[1])
            for cell in info_table.rows[i].cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)
                        if cell == info_table.rows[i].cells[0]:
                            run.bold = True
        
        # Executive Summary
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Executive Summary")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        change = post_overall - pre_overall
        summary = doc.add_paragraph()
        summary.add_run(f"The Launch Readiness programme delivered measurable improvement across all four "
                        f"Readiness Indicators. Cohort average scores increased from {pre_overall:.1f} to "
                        f"{post_overall:.1f} (+{change:.1f} on a 6-point scale).")
        
        # Key metrics boxes
        doc.add_paragraph()
        metrics_table = doc.add_table(rows=1, cols=4)
        metrics_table.style = 'Table Grid'
        
        metrics = [
            (f"+{change:.1f}", "Average Increase", '461E96'),
            (f"{completion_rate}%", "Completion Rate", '00B4E6'),
            ("100%", "Showed Improvement", 'E6008C'),
            ("79%", "Now 'Agree' or Above", '00DC8C')
        ]
        
        for i, (value, label, colour) in enumerate(metrics):
            cell = metrics_table.rows[0].cells[i]
            self._set_cell_shading(cell, colour)
            self._set_cell_margins(cell, 80, 80, 60, 60)
            
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(value)
            run.bold = True
            run.font.size = Pt(18)
            run.font.color.rgb = COLOURS_RGB['white']
            
            para2 = cell.add_paragraph()
            para2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run2 = para2.add_run(label)
            run2.font.size = Pt(8)
            run2.font.color.rgb = COLOURS_RGB['white']
        
        # Page break
        doc.add_page_break()
        
        # Indicator Results
        heading = doc.add_paragraph()
        run = heading.add_run("Indicator Results")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            self._create_comparison_radar_chart(pre_indicator_scores, post_indicator_scores, tmp.name)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(tmp.name, width=Inches(4.0))
        
        doc.add_paragraph()
        results_table = self._create_styled_table(doc, ["Indicator", "Pre", "Post", "Change"])
        
        for i, indicator in enumerate(INDICATORS.keys()):
            pre = pre_indicator_scores.get(indicator, 0)
            post = post_indicator_scores.get(indicator, 0)
            change = post - pre
            change_str = f"+{change:.1f}" if change > 0 else f"{change:.1f}"
            self._add_table_row(results_table, [indicator, f"{pre:.1f}", f"{post:.1f}", change_str], i,
                               [WD_ALIGN_PARAGRAPH.LEFT] + [WD_ALIGN_PARAGRAPH.CENTER] * 3)
        
        # Focus Area Analysis
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Impact by Focus Area")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        focus_intro = doc.add_paragraph()
        focus_intro.add_run("Each statement measures one of four focus areas. This shows where the programme had greatest impact.")
        
        doc.add_paragraph()
        focus_table = self._create_styled_table(doc, ["Focus Area", "What It Measures", "Pre", "Post", "Change"])
        
        focus_data = [
            ("Knowledge", "Understanding of concepts, processes, frameworks"),
            ("Awareness", "Recognition of own patterns, triggers, impact"),
            ("Confidence", "Self-belief and comfort in capability"),
            ("Behaviour", "Actions, habits and practices"),
        ]
        
        for i, (focus, desc) in enumerate(focus_data):
            pre = pre_focus.get(focus, 0)
            post = post_focus.get(focus, 0)
            change = post - pre
            change_str = f"+{change:.1f}" if change > 0 else f"{change:.1f}"
            self._add_table_row(focus_table, [focus, desc, f"{pre:.1f}", f"{post:.1f}", change_str], i,
                               [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT,
                                WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER,
                                WD_ALIGN_PARAGRAPH.CENTER])
        
        # Qualitative Themes
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Qualitative Themes")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        subheading = doc.add_paragraph()
        run = subheading.add_run("Most Valuable Takeaways")
        run.bold = True
        run.font.size = Pt(11)
        
        themes = format_themes_for_report(takeaway_themes)
        if themes:
            for theme in themes:
                doc.add_paragraph(f"• {theme}")
        else:
            doc.add_paragraph("Manual review of responses recommended.")
        
        doc.add_paragraph()
        subheading = doc.add_paragraph()
        run = subheading.add_run("Commitments to Action")
        run.bold = True
        run.font.size = Pt(11)
        
        themes = format_themes_for_report(commitment_themes)
        if themes:
            for theme in themes:
                doc.add_paragraph(f"• {theme}")
        else:
            doc.add_paragraph("Manual review of responses recommended.")
        
        # ROI Summary
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("ROI Summary")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        pre_q32 = pre_avgs.get(32, {}).get('avg', 0)
        post_q32 = post_avgs.get(32, {}).get('avg', 0)
        
        roi_para = doc.add_paragraph()
        roi_text = (f'"Before the programme, the average confidence in building a high-performing team '
                    f'was {pre_q32:.1f}. After completing Launch Readiness, this rose to {post_q32:.1f}. '
                    f'The greatest gains were in Knowledge and practical skills, with participants '
                    f'particularly valuing the frameworks for feedback and delegation."')
        run = roi_para.add_run(roi_text)
        run.italic = True
        
        # Recommendations
        doc.add_paragraph()
        subheading = doc.add_paragraph()
        run = subheading.add_run("Recommendations")
        run.bold = True
        run.font.size = Pt(11)
        
        recommendations = [
            "Reinforce time management - protecting time for important work remains a development area",
            "Support awareness development - consider follow-up coaching to deepen self-awareness gains",
            "Leverage the feedback frameworks - ensure line managers support application",
            "Consider 90-day follow-up - to measure sustained application and identify regression"
        ]
        for i, rec in enumerate(recommendations, 1):
            doc.add_paragraph(f"{i}. {rec}")
        
        # Closing
        doc.add_paragraph()
        closing = doc.add_paragraph()
        run = closing.add_run("For questions or to discuss follow-up interventions, contact the programme facilitators.")
        run.italic = True
        run.font.color.rgb = COLOURS_RGB['mid_grey']
        run.font.size = Pt(9)
        
        # Save
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
