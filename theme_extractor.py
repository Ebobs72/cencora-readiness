#!/usr/bin/env python3
"""
Theme extraction using Claude API for qualitative analysis.

Extracts themes from open-ended responses for cohort Impact reports.
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
