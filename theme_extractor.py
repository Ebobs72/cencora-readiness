#!/usr/bin/env python3
"""
Theme extraction using Claude API for qualitative analysis.

Extracts themes from open-ended responses for cohort Impact reports.
Enhanced to combine quantitative score data with qualitative responses
for integrated cohort-level insights.
"""

import json
import os

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False


class ThemeExtractor:
    def __init__(self):
        self.client = None
        self.api_key = None
        
        # Try Streamlit secrets first
        try:
            import streamlit as st
            self.api_key = st.secrets.get("anthropic", {}).get("api_key")
        except:
            pass
        
        # Fall back to environment variable
        if not self.api_key:
            self.api_key = os.environ.get("ANTHROPIC_API_KEY")
        
        if self.api_key and ANTHROPIC_AVAILABLE:
            self.client = anthropic.Anthropic(api_key=self.api_key)
    
    def is_available(self) -> bool:
        """Check if theme extraction is available."""
        return self.client is not None
    
    def extract_themes(self, responses: list, question_context: str, max_themes: int = 5) -> dict:
        """
        Extract themes from a list of responses.
        
        Args:
            responses: List of response texts
            question_context: Description of what question was asked
            max_themes: Maximum number of themes to extract
        
        Returns:
            dict with 'themes' list containing theme objects with 'theme', 'count', 'example'
        """
        if not self.client:
            return self._fallback_extraction(responses)
        
        if not responses:
            return {'themes': [], 'total_responses': 0}
        
        # Filter out empty responses
        valid_responses = [r for r in responses if r and r.strip()]
        
        if not valid_responses:
            return {'themes': [], 'total_responses': 0}
        
        prompt = f"""Analyse these {len(valid_responses)} responses to the question about "{question_context}".

Identify the {max_themes} most common themes, with approximate frequency counts.

Guidelines:
- Look for recurring ideas, tools, concepts, or commitments mentioned
- Count how many responses touch on each theme (a response can contribute to multiple themes)
- Select a brief, representative quote for each theme
- Be specific - "feedback models" is better than "communication skills"

Responses:
{chr(10).join(f'- "{r}"' for r in valid_responses)}

Return ONLY valid JSON in this exact format, no other text:
{{"themes": [{{"theme": "Brief theme description", "count": N, "example": "Short representative quote"}}]}}
"""
        
        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1000,
                messages=[{"role": "user", "content": prompt}]
            )
            
            result_text = response.content[0].text.strip()
            
            # Try to parse JSON
            try:
                result = json.loads(result_text)
            except json.JSONDecodeError:
                # Try to extract JSON from response
                import re
                json_match = re.search(r'\{.*\}', result_text, re.DOTALL)
                if json_match:
                    result = json.loads(json_match.group())
                else:
                    return self._fallback_extraction(valid_responses)
            
            result['total_responses'] = len(valid_responses)
            return result
            
        except Exception as e:
            print(f"Theme extraction error: {e}")
            return self._fallback_extraction(valid_responses)
    
    def _fallback_extraction(self, responses: list) -> dict:
        """Simple fallback when API is unavailable - just return response count."""
        return {
            'themes': [],
            'total_responses': len(responses) if responses else 0,
            'note': 'AI theme extraction unavailable - manual review recommended'
        }
    
    def extract_takeaways(self, responses: list) -> dict:
        """Extract themes from 'most valuable takeaway' responses."""
        return self.extract_themes(
            responses, 
            "most valuable takeaway from the programme",
            max_themes=5
        )
    
    def extract_commitments(self, responses: list) -> dict:
        """Extract themes from 'what will you do differently' responses."""
        return self.extract_themes(
            responses,
            "what participants will do differently",
            max_themes=5
        )
    
    def extract_concern_reflections(self, pre_concerns: list, post_reflections: list) -> dict:
        """
        Analyse how concerns were addressed.
        
        Args:
            pre_concerns: List of original concerns
            post_reflections: List of post-programme reflections on those concerns
        """
        if not self.client:
            return self._fallback_extraction(post_reflections)
        
        if not post_reflections:
            return {'themes': [], 'total_responses': 0}
        
        prompt = f"""Analyse how participants' pre-programme concerns were addressed.

Pre-programme concerns included themes like:
{chr(10).join(f'- "{c}"' for c in pre_concerns[:10] if c)}

Post-programme reflections on those concerns:
{chr(10).join(f'- "{r}"' for r in post_reflections if r and r.strip())}

Identify 3-5 themes in how concerns were resolved or addressed. Focus on:
- What helped alleviate concerns
- How perceptions changed
- What practical tools or insights made the difference

Return ONLY valid JSON:
{{"themes": [{{"theme": "How concerns were addressed", "count": N, "example": "Quote showing resolution"}}]}}
"""
        
        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=1000,
                messages=[{"role": "user", "content": prompt}]
            )
            
            result_text = response.content[0].text.strip()
            
            try:
                result = json.loads(result_text)
            except json.JSONDecodeError:
                import re
                json_match = re.search(r'\{.*\}', result_text, re.DOTALL)
                if json_match:
                    result = json.loads(json_match.group())
                else:
                    return self._fallback_extraction(post_reflections)
            
            result['total_responses'] = len([r for r in post_reflections if r and r.strip()])
            return result
            
        except Exception as e:
            print(f"Theme extraction error: {e}")
            return self._fallback_extraction(post_reflections)
    
    def extract_cohort_insights(self, score_data: dict, open_responses: dict) -> dict:
        """
        Generate integrated cohort insights by feeding both quantitative scores
        and qualitative responses to Claude.
        
        Args:
            score_data: dict containing:
                - n_participants: int
                - pre_overall: float
                - post_overall: float
                - indicator_scores: list of dicts with 'name', 'pre', 'post', 'change'
                - focus_scores: list of dicts with 'name', 'pre', 'post', 'change'
                - top_growth_items: list of dicts with 'num', 'text', 'pre_avg', 'post_avg', 'change'
                - lowest_post_items: list of dicts with 'num', 'text', 'post_avg'
                - pct_improved: float (% of participants who improved overall)
                - pct_agree_or_above: float (% of post items scoring >= 5)
            open_responses: dict containing:
                - takeaways: list of response strings
                - commitments: list of response strings
                - concerns_pre: list of response strings
                - concerns_post: list of response strings
        
        Returns:
            dict with:
                - executive_narrative: str (2-3 paragraph synthesis)
                - roi_narrative: str (1 paragraph for ROI section)
                - recommendations: list of str (4-5 data-driven recommendations)
                - takeaway_themes: list of dicts with 'theme', 'count', 'example'
                - commitment_themes: list of dicts with 'theme', 'count', 'example'
        """
        if not self.client:
            return self._fallback_cohort_insights(score_data, open_responses)
        
        # Build the comprehensive prompt
        n = score_data['n_participants']
        pre_o = score_data['pre_overall']
        post_o = score_data['post_overall']
        change_o = post_o - pre_o
        
        # Format indicator scores
        indicator_lines = []
        for ind in score_data['indicator_scores']:
            indicator_lines.append(
                f"  - {ind['name']}: Pre {ind['pre']:.1f} → Post {ind['post']:.1f} (change: {ind['change']:+.1f})"
            )
        
        # Format focus scores
        focus_lines = []
        for foc in score_data['focus_scores']:
            focus_lines.append(
                f"  - {foc['name']}: Pre {foc['pre']:.1f} → Post {foc['post']:.1f} (change: {foc['change']:+.1f})"
            )
        
        # Format top growth items
        growth_lines = []
        for item in score_data['top_growth_items'][:5]:
            growth_lines.append(
                f"  - Item {item['num']}: \"{item['text']}\" — Pre {item['pre_avg']:.1f} → Post {item['post_avg']:.1f} ({item['change']:+.1f})"
            )
        
        # Format lowest post items
        low_lines = []
        for item in score_data['lowest_post_items'][:5]:
            low_lines.append(
                f"  - Item {item['num']}: \"{item['text']}\" — Post avg {item['post_avg']:.1f}"
            )
        
        # Format open responses
        takeaway_text = chr(10).join(f'  - "{r}"' for r in open_responses.get('takeaways', []) if r and r.strip())
        commitment_text = chr(10).join(f'  - "{r}"' for r in open_responses.get('commitments', []) if r and r.strip())
        concern_pre_text = chr(10).join(f'  - "{r}"' for r in open_responses.get('concerns_pre', []) if r and r.strip())
        concern_post_text = chr(10).join(f'  - "{r}"' for r in open_responses.get('concerns_post', []) if r and r.strip())
        
        prompt = f"""You are analysing results from a leadership development programme called "Launch Readiness" 
designed for new leaders and managers. {n} participants completed both Pre and Post assessments.

=== QUANTITATIVE DATA ===

Overall: Pre {pre_o:.1f} → Post {post_o:.1f} ({change_o:+.1f} on a 6-point scale)
{score_data.get('pct_improved', 0):.0f}% of participants showed overall improvement
{score_data.get('pct_agree_or_above', 0):.0f}% of post-programme item scores are now at "Agree" (5) or above

Indicator Scores:
{chr(10).join(indicator_lines)}

Focus Area Scores (Knowledge/Awareness/Confidence/Behaviour):
{chr(10).join(focus_lines)}

Top 5 Growth Items (biggest cohort-average increase):
{chr(10).join(growth_lines)}

5 Items Still Scoring Lowest Post-Programme (ongoing development needs):
{chr(10).join(low_lines)}

=== QUALITATIVE DATA ===

"What was your most valuable takeaway?"
{takeaway_text or "  No responses"}

"What will you do differently?"
{commitment_text or "  No responses"}

Pre-programme concerns:
{concern_pre_text or "  No responses"}

Post-programme reflections on concerns:
{concern_post_text or "  No responses"}

=== INSTRUCTIONS ===

Based on ALL the data above, return ONLY valid JSON (no other text) in this exact format:

{{
  "executive_narrative": "Maximum 4-5 sentences. Be direct and punchy — lead with the headline result, then the strongest insight, then the one area to watch. Reference specific scores but don't list every indicator. Connect one quantitative pattern to what participants actually said. Write for a time-poor senior stakeholder who wants the story in 30 seconds.",
  
  "roi_narrative": "One paragraph suitable for an ROI section. Reference specific before/after scores, the focus areas with biggest gains, and what participants said they will do differently. Write as if quoting from a programme evaluation — authoritative but accessible.",
  
  "recommendations": [
    "4-5 specific, data-driven recommendations. Each should reference actual score patterns or qualitative themes. Format: 'Recommendation title — brief explanation referencing the data'"
  ],
  
  "takeaway_themes": [
    {{"theme": "Specific theme description", "count": N, "example": "Brief representative quote"}}
  ],
  
  "commitment_themes": [
    {{"theme": "Specific theme description", "count": N, "example": "Brief representative quote"}}
  ]
}}
"""
        
        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=3000,
                messages=[{"role": "user", "content": prompt}]
            )
            
            result_text = response.content[0].text.strip()
            
            try:
                result = json.loads(result_text)
            except json.JSONDecodeError:
                import re
                json_match = re.search(r'\{.*\}', result_text, re.DOTALL)
                if json_match:
                    result = json.loads(json_match.group())
                else:
                    return self._fallback_cohort_insights(score_data, open_responses)
            
            # Ensure all expected keys exist
            result.setdefault('executive_narrative', '')
            result.setdefault('roi_narrative', '')
            result.setdefault('recommendations', [])
            result.setdefault('takeaway_themes', [])
            result.setdefault('commitment_themes', [])
            result['total_responses'] = n
            
            return result
            
        except Exception as e:
            print(f"Cohort insights extraction error: {e}")
            return self._fallback_cohort_insights(score_data, open_responses)
    
    def _fallback_cohort_insights(self, score_data: dict, open_responses: dict) -> dict:
        """Fallback when API is unavailable — returns structured data for manual population."""
        change = score_data['post_overall'] - score_data['pre_overall']
        
        # Find biggest growth indicator
        best_ind = max(score_data['indicator_scores'], key=lambda x: x['change'])
        # Find lowest post indicator
        weakest_ind = min(score_data['indicator_scores'], key=lambda x: x['post'])
        # Find biggest growth focus
        best_focus = max(score_data['focus_scores'], key=lambda x: x['change'])
        
        executive_narrative = (
            f"Cohort scores rose from {score_data['pre_overall']:.1f} to "
            f"{score_data['post_overall']:.1f} (+{change:.1f}), with "
            f"{score_data.get('pct_improved', 0):.0f}% of participants improving overall. "
            f"Strongest gains came in {best_ind['name']} ({best_ind['change']:+.1f}). "
            f"{weakest_ind['name']} ({weakest_ind['post']:.1f} post-programme) remains the "
            f"key area for continued development."
        )
        
        roi_narrative = (
            f"Before the programme, the average readiness score was {score_data['pre_overall']:.1f}. "
            f"After completing Launch Readiness, this rose to {score_data['post_overall']:.1f}. "
            f"{score_data.get('pct_agree_or_above', 0):.0f}% of post-programme item scores now sit at "
            f"'Agree' or above. The greatest gains were in {best_focus['name'].lower()} "
            f"({best_focus['change']:+.1f}), indicating participants developed stronger understanding "
            f"of practical frameworks and tools."
        )
        
        recommendations = [
            f"Reinforce {weakest_ind['name']} — this indicator scored lowest post-programme ({weakest_ind['post']:.1f}) and may benefit from targeted follow-up",
            f"Build on {best_ind['name']} momentum — the strongest growth area ({best_ind['change']:+.1f}) suggests high receptivity to continued development",
            "Consider 90-day follow-up assessment — to measure sustained application and identify regression",
            "Ensure line manager support — sustained behaviour change requires reinforcement in the workplace",
        ]
        
        return {
            'executive_narrative': executive_narrative,
            'roi_narrative': roi_narrative,
            'recommendations': recommendations,
            'takeaway_themes': [],
            'commitment_themes': [],
            'total_responses': score_data['n_participants'],
            'note': 'AI synthesis unavailable — using data-driven fallback'
        }


def format_themes_for_report(theme_data: dict) -> list:
    """Format extracted themes for inclusion in a report."""
    if not theme_data or not theme_data.get('themes'):
        return []
    
    formatted = []
    for t in theme_data['themes']:
        theme_text = t.get('theme', '')
        count = t.get('count', 0)
        total = theme_data.get('total_responses', 0)
        
        if count and total:
            formatted.append(f"{theme_text} (mentioned by {count}/{total} participants)")
        else:
            formatted.append(theme_text)
    
    return formatted


def format_insight_themes(themes: list, total: int) -> list:
    """Format themes from extract_cohort_insights for inclusion in a report."""
    if not themes:
        return []
    
    formatted = []
    for t in themes:
        theme_text = t.get('theme', '')
        count = t.get('count', 0)
        
        if count and total:
            formatted.append(f"{theme_text} (mentioned by {count}/{total} participants)")
        else:
            formatted.append(theme_text)
    
    return formatted
