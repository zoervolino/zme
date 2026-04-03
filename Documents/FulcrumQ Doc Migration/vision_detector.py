# vision_detector.py

import base64
import requests
from pathlib import Path

class VisionDetector:
    """Calls agent relay to detect slide type and properties."""
    
    def __init__(self, relay_config: dict):
        self.agent_slug = relay_config['agent_slug']
        self.relay_url = relay_config['relay_url']
        self.api_key = relay_config['api_key']
    
    def detect(self, image_path: Path, slide_num: int) -> dict:
        """
        Send image to agent relay for analysis.
        Returns: {slide_type, title, subtitle, confidence, colors, ...}
        """
        # Read image
        with open(image_path, 'rb') as f:
            image_data = base64.b64encode(f.read()).decode('utf-8')
        
        # Call agent
        prompt = f"""Analyze slide {slide_num}. Return JSON:
{{
  "slide_type": "cover|divider|content|ending|blank",
  "title": "detected title or null",
  "subtitle": "detected subtitle or null",
  "confidence": 0.0-1.0,
  "dominant_colors": ["#RRGGBB", ...],
  "has_table": true/false,
  "has_image": true/false,
  "text_density": "low|medium|high"
}}
"""
        
        response = requests.post(
            f"{self.relay_url}/agents/{self.agent_slug}/invoke",
            headers={'Authorization': f'Bearer {self.api_key}'},
            json={
                'prompt': prompt,
                'image': image_data,
            }
        )
        
        result = response.json()
        return result.get('analysis', {})
