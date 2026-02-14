#!/usr/bin/env python3
"""
Report generator for Launch Readiness assessments.
All reports use Cencora branding and house style.
"""

import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Twips, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

from framework import (
    INDICATORS, INDICATOR_DESCRIPTIONS, INDICATOR_COLOURS,
    ITEMS, OPEN_QUESTIONS_PRE, OPEN_QUESTIONS_POST,
    FOCUS_TAGS, get_items_by_focus
)
from theme_extractor import ThemeExtractor, format_themes_for_report

COLOURS_HEX = {
    'purple': '#461E96', 'cyan': '#00B4E6', 'magenta': '#E6008C',
    'green': '#00DC8C', 'mid_grey': '#6E6E6E', 'success_green': '#2E7D32'
}

COLOURS_RGB = {
    'purple': RGBColor(0x46, 0x1E, 0x96), 'cyan': RGBColor(0x00, 0xB4, 0xE6),
    'magenta': RGBColor(0xE6, 0x00, 0x8C), 'green': RGBColor(0x00, 0xDC, 0x8C),
    'mid_grey': RGBColor(0x6E, 0x6E, 0x6E), 'success_green': RGBColor(0x2E, 0x7D, 0x32),
    'white': RGBColor(0xFF, 0xFF, 0xFF)
}

INDICATOR_COLOUR_MAP = {
    'Self-Readiness': 'purple', 'Practical Readiness': 'cyan',
    'Professional Readiness': 'magenta', 'Team Readiness': 'green'
}

# Column widths for 5-column item tables (in inches) - matching sample
COL_WIDTHS_5 = [0.3, 4.5, 0.9, 0.6, 0.5]


class ReportGenerator:
    def __init__(self, db):
        self.db = db
        self.theme_extractor = ThemeExtractor()
    
    def _find_logo_path(self):
        """Find the logo file in various possible locations."""
        possible_paths = [
            Path(__file__).parent / 'assets' / 'cencora_logo.png',
            Path('/mount/src/cencora-readiness/assets/cencora_logo.png'),
            Path('assets/cencora_logo.png'),
            Path('./assets/cencora_logo.png'),
        ]
        for path in possible_paths:
            if path.exists():
                return path
        return None
    
    def _create_radar_chart(self, scores: dict, output_path: str):
        """Create a 4-axis radar chart with proper sizing and labels."""
        indicators = list(INDICATORS.keys())
        values = [scores.get(ind, 0) for ind in indicators]
        
        fig = plt.figure(figsize=(8, 8))
        ax = fig.add_subplot(111, polar=True)
        fig.patch.set_facecolor('white')
        
        # KEY: Use evenly spaced angles around the full circle
        num_vars = 4
        angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
        
        # Close the polygon
        values_closed = values + [values[0]]
        angles_closed = angles + [angles[0]]
        
        # Set theta offset so Self-Readiness is at the top, clockwise direction
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        
        # Grid
        ax.set_ylim(0, 6)
        ax.set_yticks([1, 2, 3, 4, 5, 6])
        ax.set_yticklabels(['1', '2', '3', '4', '5', '6'], fontsize=10, color='#888888')
        ax.yaxis.grid(True, color='#E0E0E0', linewidth=1)
        
        # Data polygon
        ax.fill(angles_closed, values_closed, color=COLOURS_HEX['purple'], alpha=0.25)
        ax.plot(angles_closed, values_closed, color=COLOURS_HEX['purple'], linewidth=3)
        
        # Data points with indicator colours
        indicator_colours = [COLOURS_HEX['purple'], COLOURS_HEX['cyan'], 
                           COLOURS_HEX['magenta'], COLOURS_HEX['green']]
        for angle, value, colour in zip(angles, values, indicator_colours):
            ax.scatter(angle, value, color=colour, s=200, zorder=5, edgecolors='white', linewidths=2)
        
        # Remove default labels
        ax.set_xticks(angles)
        ax.set_xticklabels([])
        
        # Add custom labels
        label_distance = 7.2
        for i, (ind, colour) in enumerate(zip(indicators, indicator_colours)):
            angle = angles[i]
            if i == 0:  # Top
                ha, va = 'center', 'bottom'
            elif i == 1:  # Right
                ha, va = 'left', 'center'
            elif i == 2:  # Bottom
                ha, va = 'center', 'top'
            else:  # Left
                ha, va = 'right', 'center'
            ax.text(angle, label_distance, ind, ha=ha, va=va, fontsize=18, fontweight='bold', color=colour)
        
        ax.spines['polar'].set_visible(False)
        
        plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white', pad_inches=0.3)
        plt.close()
    
    def _create_comparison_radar_chart(self, pre_scores: dict, post_scores: dict, output_path: str):
        """Create a comparison radar chart (pre dashed grey, post solid green)."""
        indicators = list(INDICATORS.keys())
        pre_values = [pre_scores.get(ind, 0) for ind in indicators]
        post_values = [post_scores.get(ind, 0) for ind in indicators]
        
        fig = plt.figure(figsize=(8, 8))
        ax = fig.add_subplot(111, polar=True)
        fig.patch.set_facecolor('white')
        
        # KEY: Use evenly spaced angles around the full circle
        num_vars = 4
        angles = np.linspace(0, 2 * np.pi, num_vars, endpoint=False).tolist()
        
        # Close the polygons
        pre_closed = pre_values + [pre_values[0]]
        post_closed = post_values + [post_values[0]]
        angles_closed = angles + [angles[0]]
        
        # Set theta offset so Self-Readiness is at the top, clockwise direction
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        
        ax.set_ylim(0, 6)
        ax.set_yticks([1, 2, 3, 4, 5, 6])
        ax.set_yticklabels(['1', '2', '3', '4', '5', '6'], fontsize=10, color='#888888')
        ax.yaxis.grid(True, color='#E0E0E0', linewidth=1)
        
        # PRE polygon (dashed, grey)
        ax.fill(angles_closed, pre_closed, color='#999999', alpha=0.1)
        ax.plot(angles_closed, pre_closed, color='#999999', linewidth=2, linestyle='--')
        
        # POST polygon (solid, green)
        ax.fill(angles_closed, post_closed, color=COLOURS_HEX['green'], alpha=0.25)
        ax.plot(angles_closed, post_closed, color=COLOURS_HEX['green'], linewidth=3)
        
        indicator_colours = [COLOURS_HEX['purple'], COLOURS_HEX['cyan'], 
                           COLOURS_HEX['magenta'], COLOURS_HEX['green']]
        for angle, pre_val, post_val, colour in zip(angles, pre_values, post_values, indicator_colours):
            ax.scatter(angle, pre_val, color='#999999', s=80, zorder=4, edgecolors='white', linewidths=1)
            ax.scatter(angle, post_val, color=colour, s=200, zorder=5, edgecolors='white', linewidths=2)
        
        ax.set_xticks(angles)
        ax.set_xticklabels([])
        
        # Add custom labels
        label_distance = 7.2
        for i, (ind, colour) in enumerate(zip(indicators, indicator_colours)):
            angle = angles[i]
            if i == 0:  # Top
                ha, va = 'center', 'bottom'
            elif i == 1:  # Right
                ha, va = 'left', 'center'
            elif i == 2:  # Bottom
                ha, va = 'center', 'top'
            else:  # Left
                ha, va = 'right', 'center'
            ax.text(angle, label_distance, ind, ha=ha, va=va, fontsize=18, fontweight='bold', color=colour)
        
        # Legend
        from matplotlib.lines import Line2D
        legend_elements = [
            Line2D([0], [0], color='#999999', linestyle='--', linewidth=2, label='Pre-Programme'),
            Line2D([0], [0], color=COLOURS_HEX['green'], linewidth=3, label='Post-Programme')
        ]
        ax.legend(handles=legend_elements, loc='lower left', bbox_to_anchor=(-0.15, -0.15),
                 fontsize=11, frameon=False)
        
        ax.spines['polar'].set_visible(False)
        plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white', pad_inches=0.3)
        plt.close()
    
    def _create_bar_chart(self, score: float, colour_hex: str, output_path: str):
        """Create a horizontal bar chart for a single score."""
        fig, ax = plt.subplots(figsize=(1.2, 0.3))
        fig.patch.set_facecolor('white')
        ax.barh(0, 6, color='#E8E8E8', height=0.6)
        ax.barh(0, score, color=colour_hex, height=0.6)
        ax.set_xlim(0, 6)
        ax.set_ylim(-0.5, 0.5)
        ax.axis('off')
        plt.savefig(output_path, dpi=100, bbox_inches='tight', facecolor='white', pad_inches=0.01)
        plt.close()
    
    def _create_comparison_bar_chart(self, pre_score: float, post_score: float, colour_hex: str, output_path: str):
        """Create a comparison bar (pre dashed outline, post filled)."""
        fig, ax = plt.subplots(figsize=(1.2, 0.3))
        fig.patch.set_facecolor('white')
        ax.barh(0, 6, color='#E8E8E8', height=0.6)
        ax.barh(0, post_score, color=colour_hex, height=0.6)
        ax.barh(0, pre_score, color='none', edgecolor='#666666', linewidth=1.5, linestyle='--', height=0.6)
        ax.set_xlim(0, 6)
        ax.set_ylim(-0.5, 0.5)
        ax.axis('off')
        plt.savefig(output_path, dpi=100, bbox_inches='tight', facecolor='white', pad_inches=0.01)
        plt.close()
    
    def _set_cell_shading(self, cell, colour_hex: str):
        """Set cell background colour."""
        colour_hex = colour_hex.replace('#', '')
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), colour_hex)
        cell._tc.get_or_add_tcPr().append(shading)
    
    def _set_cell_margins(self, cell, top=40, bottom=40, left=80, right=80):
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
    
    def _set_column_width(self, column, width_inches):
        """Set column width."""
        for cell in column.cells:
            cell.width = Inches(width_inches)
    
    def _add_logo_header(self, doc):
        """Add Cencora logo to document header."""
        logo_path = self._find_logo_path()
        if logo_path:
            try:
                section = doc.sections[0]
                header = section.header
                header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                run = header_para.add_run()
                run.add_picture(str(logo_path), width=Inches(1.2))
            except Exception as e:
                print(f"Logo error: {e}")
    
    def _create_info_table(self, doc, rows_data):
        """Create the participant info table."""
        table = doc.add_table(rows=len(rows_data), cols=2)
        table.style = 'Table Grid'
        
        for i, (label, value) in enumerate(rows_data):
            table.rows[i].cells[0].text = label
            table.rows[i].cells[1].text = str(value)
            self._set_cell_shading(table.rows[i].cells[0], 'F5F5F5')
            self._set_cell_margins(table.rows[i].cells[0])
            self._set_cell_margins(table.rows[i].cells[1])
            
            # Set column widths
            table.rows[i].cells[0].width = Inches(1.45)
            table.rows[i].cells[1].width = Inches(4.85)
            
            for cell in table.rows[i].cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.name = 'Arial'
                        if cell == table.rows[i].cells[0]:
                            run.bold = True
        return table
    
    def _create_summary_table(self, doc, indicator_scores, overall_score):
        """Create the summary indicator/score table."""
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        
        # Header row
        table.rows[0].cells[0].text = "Indicator"
        table.rows[0].cells[1].text = "Score"
        for cell in table.rows[0].cells:
            self._set_cell_shading(cell, '461E96')
            self._set_cell_margins(cell)
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.bold = True
                    run.font.color.rgb = COLOURS_RGB['white']
                    run.font.size = Pt(10)
                    run.font.name = 'Arial'
        
        # Set column widths
        table.rows[0].cells[0].width = Inches(5.85)
        table.rows[0].cells[1].width = Inches(1.39)
        
        # Data rows
        for i, (ind, score) in enumerate(indicator_scores.items()):
            row = table.add_row()
            row.cells[0].text = ind
            row.cells[1].text = f"{score:.1f}"
            bg = 'FFFFFF' if i % 2 == 0 else 'FDF6E3'
            for j, cell in enumerate(row.cells):
                self._set_cell_shading(cell, bg)
                self._set_cell_margins(cell)
                cell.width = Inches(5.85) if j == 0 else Inches(1.39)
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if j == 1 else WD_ALIGN_PARAGRAPH.LEFT
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.name = 'Arial'
        
        # Overall row
        row = table.add_row()
        row.cells[0].text = "OVERALL"
        row.cells[1].text = f"{overall_score:.1f}"
        for j, cell in enumerate(row.cells):
            self._set_cell_shading(cell, 'F5F5F5')
            self._set_cell_margins(cell)
            cell.width = Inches(5.85) if j == 0 else Inches(1.39)
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER if j == 1 else WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(10)
                    run.font.name = 'Arial'
        
        return table
    
    def _create_items_table(self, doc, indicator, start, end, ratings, colour_hex):
        """Create a detailed items table for an indicator."""
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        
        # Header
        headers = ["#", "Statement", "Focus", "", "Score"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
            self._set_cell_shading(table.rows[0].cells[i], colour_hex.replace('#', ''))
            self._set_cell_margins(table.rows[0].cells[i])
            table.rows[0].cells[i].width = Inches(COL_WIDTHS_5[i])
            for para in table.rows[0].cells[i].paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.bold = True
                    run.font.color.rgb = COLOURS_RGB['white']
                    run.font.size = Pt(9)
                    run.font.name = 'Arial'
        
        # Data rows
        for idx, item_num in enumerate(range(start, end + 1)):
            item = ITEMS[item_num]
            score = ratings.get(item_num, 0)
            row = table.add_row()
            bg = 'FFFFFF' if idx % 2 == 0 else 'FDF6E3'
            
            # Create bar chart
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                self._create_bar_chart(score, colour_hex, tmp.name)
                bar_path = tmp.name
            
            values = [str(item_num), item['text'], item['focus'], None, str(score)]
            alignments = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, 
                         WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]
            
            for j, cell in enumerate(row.cells):
                cell.width = Inches(COL_WIDTHS_5[j])
                self._set_cell_shading(cell, bg)
                self._set_cell_margins(cell)
                
                if j == 3:  # Bar chart column
                    para = cell.paragraphs[0]
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = para.add_run()
                    run.add_picture(bar_path, width=Inches(0.7))
                else:
                    cell.text = str(values[j]) if values[j] is not None else ''
                    for para in cell.paragraphs:
                        para.alignment = alignments[j]
                        for run in para.runs:
                            run.font.size = Pt(9)
                            run.font.name = 'Arial'
        
        return table
    
    def _calculate_indicator_scores(self, ratings):
        scores = {}
        for indicator, (start, end) in INDICATORS.items():
            item_scores = [ratings.get(i, 0) for i in range(start, end + 1) if i in ratings]
            scores[indicator] = sum(item_scores) / len(item_scores) if item_scores else 0
        return scores
    
    def _calculate_overall_score(self, ratings):
        if not ratings:
            return 0
        valid_ratings = [ratings.get(i, 0) for i in range(1, 33) if i in ratings]
        return sum(valid_ratings) / len(valid_ratings) if valid_ratings else 0
    
    def generate_baseline_report(self, participant_id):
        """Generate a Baseline report (PRE assessment only)."""
        data = self.db.get_participant_data(participant_id)
        if not data or not data['pre']:
            raise ValueError("No PRE assessment data found")
        
        participant = data['participant']
        cohort = data['cohort']
        pre_ratings = data['pre']['ratings']
        pre_responses = data['pre']['open_responses']
        pre_date = data['pre']['assessment'].get('completed_at', '')[:10]
        
        indicator_scores = self._calculate_indicator_scores(pre_ratings)
        overall_score = self._calculate_overall_score(pre_ratings)
        
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        
        # Add logo
        self._add_logo_header(doc)
        
        # Title
        title = doc.add_paragraph()
        run = title.add_run("THE READINESS FRAMEWORK")
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = COLOURS_RGB['purple']
        run.font.name = 'Arial'
        
        subtitle = doc.add_paragraph()
        run = subtitle.add_run("Readiness Baseline")
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['magenta']
        run.font.name = 'Arial'
        
        # Info table
        self._create_info_table(doc, [
            ("Participant:", participant['name']),
            ("Role:", participant.get('role', 'Not specified')),
            ("Cohort:", cohort['name']),
            ("Assessment Date:", pre_date)
        ])
        
        # Introduction
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Your Starting Point")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        run.font.name = 'Arial'
        
        first_name = participant['name'].split()[0]
        intro = doc.add_paragraph()
        run = intro.add_run(f"Welcome to the Launch Readiness programme, {first_name}. "
                           f"This report captures your self-assessment before the programme begins. "
                           f"There are no right or wrong answers – this is simply a snapshot of where you see yourself today.")
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        
        # Radar chart section
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Your Readiness Profile")
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['purple']
        run.font.name = 'Arial'
        
        # Create and insert radar chart
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            self._create_radar_chart(indicator_scores, tmp.name)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(tmp.name, width=Inches(4.5))
        
        # Scale note
        scale_para = doc.add_paragraph()
        scale_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = scale_para.add_run("Scale: 1-6 (1=Strongly Disagree, 6=Strongly Agree)")
        run.italic = True
        run.font.size = Pt(9)
        run.font.color.rgb = COLOURS_RGB['mid_grey']
        run.font.name = 'Arial'
        
        doc.add_paragraph()
        
        # Summary table
        self._create_summary_table(doc, indicator_scores, overall_score)
        
        doc.add_page_break()
        
        # Detailed scores by indicator
        for indicator, (start, end) in INDICATORS.items():
            colour_name = INDICATOR_COLOUR_MAP.get(indicator, 'purple')
            colour_hex = COLOURS_HEX[colour_name]
            
            # Indicator heading
            heading = doc.add_paragraph()
            run = heading.add_run(indicator)
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = COLOURS_RGB[colour_name]
            run.font.name = 'Arial'
            
            # Description with average
            desc = doc.add_paragraph()
            run = desc.add_run(INDICATOR_DESCRIPTIONS.get(indicator, ''))
            run.italic = True
            run.font.color.rgb = COLOURS_RGB['mid_grey']
            run.font.size = Pt(9)
            run.font.name = 'Arial'
            run = desc.add_run(f"  |  Dimension Average: ")
            run.font.size = Pt(9)
            run.font.name = 'Arial'
            run = desc.add_run(f"{indicator_scores.get(indicator, 0):.1f}")
            run.bold = True
            run.font.size = Pt(9)
            run.font.name = 'Arial'
            
            # Items table
            self._create_items_table(doc, indicator, start, end, pre_ratings, colour_hex)
            doc.add_paragraph()
        
        # Overall Readiness section
        heading = doc.add_paragraph()
        run = heading.add_run("Overall Readiness")
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['purple']
        run.font.name = 'Arial'
        
        self._create_items_table(doc, "Overall", 31, 32, pre_ratings, COLOURS_HEX['purple'])
        
        doc.add_page_break()
        
        # Reflections section
        heading = doc.add_paragraph()
        run = heading.add_run("Your Reflections")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        run.font.name = 'Arial'
        
        for q_num, question in OPEN_QUESTIONS_PRE.items():
            # Question
            q_para = doc.add_paragraph()
            run = q_para.add_run(question)
            run.bold = True
            run.font.size = Pt(10)
            run.font.name = 'Arial'
            
            # Response (in italics, indented feel)
            response = pre_responses.get(q_num, "No response provided")
            r_para = doc.add_paragraph()
            run = r_para.add_run(response)
            run.italic = True
            run.font.size = Pt(10)
            run.font.name = 'Arial'
        
        doc.add_paragraph()
        
        # Closing note
        closing = doc.add_paragraph()
        run = closing.add_run("Keep this report – you'll revisit it after the programme to see how far you've come.")
        run.italic = True
        run.font.color.rgb = COLOURS_RGB['mid_grey']
        run.font.size = Pt(9)
        run.font.name = 'Arial'
        
        # Save
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    
    def generate_progress_report(self, participant_id, cohort_id):
        """Generate a Progress report (PRE vs POST comparison)."""
        data = self.db.get_participant_data(participant_id)
        if not data or not data['pre'] or not data['post']:
            raise ValueError("Both PRE and POST assessments required")
        
        participant = data['participant']
        cohort = data['cohort']
        pre_ratings = data['pre']['ratings']
        post_ratings = data['post']['ratings']
        pre_responses = data['pre']['open_responses']
        post_responses = data['post']['open_responses']
        pre_date = data['pre']['assessment'].get('completed_at', '')[:10]
        post_date = data['post']['assessment'].get('completed_at', '')[:10]
        
        pre_indicator_scores = self._calculate_indicator_scores(pre_ratings)
        post_indicator_scores = self._calculate_indicator_scores(post_ratings)
        pre_overall = self._calculate_overall_score(pre_ratings)
        post_overall = self._calculate_overall_score(post_ratings)
        
        # Cohort averages
        cohort_avgs = self.db.get_cohort_averages(cohort_id, 'POST')
        cohort_indicator_scores = {}
        for indicator, (start, end) in INDICATORS.items():
            item_avgs = [cohort_avgs.get(i, {}).get('avg', 0) for i in range(start, end + 1)]
            cohort_indicator_scores[indicator] = sum(item_avgs) / len(item_avgs) if item_avgs else 0
        cohort_overall = sum(cohort_avgs.get(i, {}).get('avg', 0) for i in range(1, 33)) / 32
        
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        
        self._add_logo_header(doc)
        
        # Title
        title = doc.add_paragraph()
        run = title.add_run("THE READINESS FRAMEWORK")
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        subtitle = doc.add_paragraph()
        run = subtitle.add_run("Readiness Progress")
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['magenta']
        
        # Info table
        self._create_info_table(doc, [
            ("Participant:", participant['name']),
            ("Role:", participant.get('role', 'Not specified')),
            ("Pre-Assessment:", pre_date),
            ("Post-Assessment:", post_date)
        ])
        
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Your Progress")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        change = post_overall - pre_overall
        first_name = participant['name'].split()[0]
        intro = doc.add_paragraph()
        intro.add_run(f"Congratulations, {first_name}! Your overall readiness improved from "
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
            run.add_picture(tmp.name, width=Inches(4.5))
        
        # Summary table with Pre/Post/Change/Cohort
        doc.add_paragraph()
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        
        headers = ["Indicator", "Pre", "Post", "Change", "Cohort"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
            self._set_cell_shading(table.rows[0].cells[i], '461E96')
            self._set_cell_margins(table.rows[0].cells[i])
            for para in table.rows[0].cells[i].paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.bold = True
                    run.font.color.rgb = COLOURS_RGB['white']
                    run.font.size = Pt(9)
        
        for idx, indicator in enumerate(INDICATORS.keys()):
            row = table.add_row()
            pre = pre_indicator_scores.get(indicator, 0)
            post = post_indicator_scores.get(indicator, 0)
            ch = post - pre
            coh = cohort_indicator_scores.get(indicator, 0)
            values = [indicator, f"{pre:.1f}", f"{post:.1f}", f"+{ch:.1f}" if ch > 0 else f"{ch:.1f}", f"{coh:.1f}"]
            bg = 'FFFFFF' if idx % 2 == 0 else 'FDF6E3'
            for j, cell in enumerate(row.cells):
                cell.text = values[j]
                self._set_cell_shading(cell, bg)
                self._set_cell_margins(cell)
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
                    for run in para.runs:
                        run.font.size = Pt(9)
        
        # Overall row
        row = table.add_row()
        ch = post_overall - pre_overall
        values = ["OVERALL", f"{pre_overall:.1f}", f"{post_overall:.1f}", 
                  f"+{ch:.1f}" if ch > 0 else f"{ch:.1f}", f"{cohort_overall:.1f}"]
        for j, cell in enumerate(row.cells):
            cell.text = values[j]
            self._set_cell_shading(cell, 'F5F5F5')
            self._set_cell_margins(cell)
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(9)
        
        doc.add_page_break()
        
        # Detailed comparison by indicator (simplified for now)
        for indicator, (start, end) in INDICATORS.items():
            colour_name = INDICATOR_COLOUR_MAP.get(indicator, 'purple')
            colour_hex = COLOURS_HEX[colour_name]
            
            heading = doc.add_paragraph()
            run = heading.add_run(indicator)
            run.bold = True
            run.font.size = Pt(12)
            run.font.color.rgb = COLOURS_RGB[colour_name]
            
            pre_avg = pre_indicator_scores.get(indicator, 0)
            post_avg = post_indicator_scores.get(indicator, 0)
            change = post_avg - pre_avg
            
            desc = doc.add_paragraph()
            run = desc.add_run(f"Pre: {pre_avg:.1f} → Post: {post_avg:.1f} ")
            run.font.size = Pt(9)
            change_run = desc.add_run(f"(+{change:.1f})" if change > 0 else f"({change:.1f})")
            change_run.bold = True
            if change > 0:
                change_run.font.color.rgb = COLOURS_RGB['success_green']
            
            # Comparison items table
            table = doc.add_table(rows=1, cols=6)
            table.style = 'Table Grid'
            headers = ["#", "Statement", "Pre", "Post", "", "Chg"]
            col_widths = [0.3, 4, 0.45, 0.45, 0.55, 0.45]
            for i, h in enumerate(headers):
                table.rows[0].cells[i].text = h
                self._set_cell_shading(table.rows[0].cells[i], colour_hex.replace('#', ''))
                self._set_cell_margins(table.rows[0].cells[i])
                table.rows[0].cells[i].width = Inches(col_widths[i])
                for para in table.rows[0].cells[i].paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.bold = True
                        run.font.color.rgb = COLOURS_RGB['white']
                        run.font.size = Pt(8)
            
            for idx, item_num in enumerate(range(start, end + 1)):
                item = ITEMS[item_num]
                pre_score = pre_ratings.get(item_num, 0)
                post_score = post_ratings.get(item_num, 0)
                item_ch = post_score - pre_score
                
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                    self._create_comparison_bar_chart(pre_score, post_score, colour_hex, tmp.name)
                    bar_path = tmp.name
                
                row = table.add_row()
                bg = 'FFFFFF' if idx % 2 == 0 else 'FDF6E3'
                values = [str(item_num), item['text'], str(pre_score), str(post_score), None, 
                         f"+{item_ch}" if item_ch > 0 else str(item_ch)]
                
                for j, cell in enumerate(row.cells):
                    cell.width = Inches(col_widths[j])
                    self._set_cell_shading(cell, bg)
                    self._set_cell_margins(cell)
                    if j == 4:
                        para = cell.paragraphs[0]
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = para.add_run()
                        run.add_picture(bar_path, width=Inches(0.6))
                    else:
                        cell.text = str(values[j]) if values[j] is not None else ''
                        for para in cell.paragraphs:
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER if j != 1 else WD_ALIGN_PARAGRAPH.LEFT
                            for run in para.runs:
                                run.font.size = Pt(8)
            
            doc.add_paragraph()
        
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
            
            if q_num == 3:
                original = pre_responses.get(3, "")
                if original:
                    orig_para = doc.add_paragraph()
                    run = orig_para.add_run(f"Your original concern: \"{original}\"")
                    run.italic = True
                    run.font.color.rgb = COLOURS_RGB['mid_grey']
                    run.font.size = Pt(9)
            
            response = post_responses.get(q_num, "No response provided")
            r_para = doc.add_paragraph()
            run = r_para.add_run(response)
            run.italic = True
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
    
    def generate_impact_report(self, cohort_id):
        """Generate an Impact report (cohort summary)."""
        cohort_data = self.db.get_cohort_data(cohort_id)
        if not cohort_data:
            raise ValueError("Cohort not found")
        
        cohort = cohort_data['cohort']
        participants = cohort_data['participants']
        complete_participants = [p for p in participants if p['pre'] and p['post']]
        
        if len(complete_participants) < 2:
            raise ValueError("Need at least 2 participants with complete data")
        
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
        
        # Theme extraction
        takeaway_responses = [p['post']['open_responses'].get(1, '') for p in complete_participants if p['post']['open_responses'].get(1)]
        commitment_responses = [p['post']['open_responses'].get(2, '') for p in complete_participants if p['post']['open_responses'].get(2)]
        
        takeaway_themes = self.theme_extractor.extract_takeaways(takeaway_responses)
        commitment_themes = self.theme_extractor.extract_commitments(commitment_responses)
        
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        
        self._add_logo_header(doc)
        
        title = doc.add_paragraph()
        run = title.add_run("THE READINESS FRAMEWORK")
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        subtitle = doc.add_paragraph()
        run = subtitle.add_run("Readiness Impact")
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['magenta']
        
        completion_rate = int(cohort_data['post_completed'] / len(participants) * 100) if participants else 0
        self._create_info_table(doc, [
            ("Programme:", cohort.get('programme', 'Launch Readiness')),
            ("Cohort:", cohort['name']),
            ("Participants:", f"{len(participants)} enrolled | {cohort_data['post_completed']} completed"),
            ("Completion Rate:", f"{completion_rate}%")
        ])
        
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Executive Summary")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        change = post_overall - pre_overall
        summary = doc.add_paragraph()
        summary.add_run(f"The programme delivered measurable improvement. Cohort average scores increased from "
                        f"{pre_overall:.1f} to {post_overall:.1f} (+{change:.1f} on a 6-point scale).")
        
        # Comparison radar
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            self._create_comparison_radar_chart(pre_indicator_scores, post_indicator_scores, tmp.name)
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run()
            run.add_picture(tmp.name, width=Inches(4.5))
        
        # Results table
        doc.add_paragraph()
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        headers = ["Indicator", "Pre", "Post", "Change"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
            self._set_cell_shading(table.rows[0].cells[i], '461E96')
            self._set_cell_margins(table.rows[0].cells[i])
            for para in table.rows[0].cells[i].paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.font.bold = True
                    run.font.color.rgb = COLOURS_RGB['white']
                    run.font.size = Pt(10)
        
        for idx, indicator in enumerate(INDICATORS.keys()):
            row = table.add_row()
            pre = pre_indicator_scores.get(indicator, 0)
            post = post_indicator_scores.get(indicator, 0)
            ch = post - pre
            values = [indicator, f"{pre:.1f}", f"{post:.1f}", f"+{ch:.1f}" if ch > 0 else f"{ch:.1f}"]
            bg = 'FFFFFF' if idx % 2 == 0 else 'FDF6E3'
            for j, cell in enumerate(row.cells):
                cell.text = values[j]
                self._set_cell_shading(cell, bg)
                self._set_cell_margins(cell)
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
                    for run in para.runs:
                        run.font.size = Pt(10)
        
        # Themes
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Qualitative Themes")
        run.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        sub = doc.add_paragraph()
        run = sub.add_run("Most Valuable Takeaways")
        run.bold = True
        themes = format_themes_for_report(takeaway_themes)
        for theme in (themes or ["Manual review recommended"]):
            doc.add_paragraph(f"• {theme}")
        
        doc.add_paragraph()
        sub = doc.add_paragraph()
        run = sub.add_run("Commitments to Action")
        run.bold = True
        themes = format_themes_for_report(commitment_themes)
        for theme in (themes or ["Manual review recommended"]):
            doc.add_paragraph(f"• {theme}")
        
        # Recommendations
        doc.add_paragraph()
        heading = doc.add_paragraph()
        run = heading.add_run("Recommendations")
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = COLOURS_RGB['purple']
        
        for i, rec in enumerate([
            "Reinforce time management – protecting time remains a development area",
            "Support awareness development through follow-up coaching",
            "Leverage feedback frameworks – ensure line managers support application",
            "Consider 90-day follow-up to measure sustained application"
        ], 1):
            doc.add_paragraph(f"{i}. {rec}")
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer
