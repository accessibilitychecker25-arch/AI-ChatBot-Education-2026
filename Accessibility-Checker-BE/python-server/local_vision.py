"""
Local AI Vision Models for Alt Text Generation (100% FREE)
Uses Hugging Face transformers to run models locally - no API costs!

Supported models:
- BLIP: Good balance of speed and quality
- GIT: More detailed descriptions
- LLAVA: Most advanced (requires more resources)
"""

import os
from typing import Optional
from pathlib import Path
import io

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("⚠️  Pillow not installed. Run: pip install pillow")

try:
    from transformers import BlipProcessor, BlipForConditionalGeneration
    from transformers import AutoProcessor, AutoModelForCausalLM
    import torch
    TRANSFORMERS_AVAILABLE = True
except ImportError:
    TRANSFORMERS_AVAILABLE = False
    print("⚠️  Transformers not installed. Run: pip install transformers torch")


class LocalVisionModel:
    """
    Local AI model for generating image descriptions
    Runs on your computer - 100% FREE with no API limits!
    """
    
    def __init__(self, model_name: str = "blip-base"):
        """
        Initialize local vision model
        
        Args:
            model_name: Model to use
                - "blip-base" (default): Fast, good quality, ~1GB
                - "blip-large": Better quality, slower, ~2GB
                - "git-base": Alternative model, ~1.5GB
        """
        self.model_name = model_name
        self.enabled = False
        self.model = None
        self.processor = None
        self.device = "cuda" if torch.cuda.is_available() else "cpu"
        
        if not TRANSFORMERS_AVAILABLE:
            print("❌ Transformers library not available")
            print("   Install with: pip install transformers torch")
            return
        
        if not PIL_AVAILABLE:
            print("❌ Pillow not available")
            print("   Install with: pip install pillow")
            return
        
        # Load model
        try:
            print(f"📥 Loading {model_name} model... (this may take a minute on first run)")
            
            if "blip" in model_name.lower():
                self._load_blip_model(model_name)
            elif "git" in model_name.lower():
                self._load_git_model()
            else:
                print(f"⚠️  Unknown model: {model_name}, defaulting to BLIP")
                self._load_blip_model("blip-base")
            
            self.enabled = True
            print(f"✅ {model_name} model loaded successfully on {self.device}")
            
        except Exception as e:
            print(f"❌ Failed to load model: {e}")
            self.enabled = False
    
    def _load_blip_model(self, model_name: str):
        """Load BLIP model (recommended for most use cases)"""
        if "large" in model_name:
            model_id = "Salesforce/blip-image-captioning-large"
        else:
            model_id = "Salesforce/blip-image-captioning-base"
        
        self.processor = BlipProcessor.from_pretrained(model_id)
        self.model = BlipForConditionalGeneration.from_pretrained(model_id)
        self.model.to(self.device)
        self.model_type = "blip"
    
    def _load_git_model(self):
        """Load GIT model (alternative to BLIP)"""
        model_id = "microsoft/git-base"
        self.processor = AutoProcessor.from_pretrained(model_id)
        self.model = AutoModelForCausalLM.from_pretrained(model_id)
        self.model.to(self.device)
        self.model_type = "git"
    
    def is_enabled(self) -> bool:
        """Check if model is loaded and ready"""
        return self.enabled and self.model is not None
    
    def generate_alt_text(
        self,
        image_data: bytes,
        shape_name: str = "",
        slide_number: int = 0,
        max_length: int = 250
    ) -> Optional[str]:
        """
        Generate alt text for an image using local AI
        
        Args:
            image_data: Raw image bytes
            shape_name: Shape name (for context)
            slide_number: Slide number (for context)
            max_length: Maximum alt text length
            
        Returns:
            Generated alt text or None if failed
        """
        if not self.is_enabled():
            return None
        
        try:
            # Convert bytes to PIL Image
            image = Image.open(io.BytesIO(image_data)).convert("RGB")
            
            # Check if image looks decorative (very small, likely a logo/icon)
            if image.size[0] < 100 and image.size[1] < 100:
                # Small image - likely decorative
                if any(hint in shape_name.lower() for hint in ["logo", "icon", "background", "border"]):
                    return "decorative"
            
            # Generate description
            if self.model_type == "blip":
                alt_text = self._generate_blip(image)
            elif self.model_type == "git":
                alt_text = self._generate_git(image)
            else:
                return None
            
            # Clean up the text
            alt_text = self._clean_alt_text(alt_text, max_length)
            
            return alt_text
            
        except Exception as e:
            print(f"Error generating alt text: {e}")
            return None
    
    def _generate_blip(self, image: Image.Image) -> str:
        """Generate caption using BLIP model"""
        # Process image
        inputs = self.processor(image, return_tensors="pt").to(self.device)
        
        # Generate caption
        with torch.no_grad():
            out = self.model.generate(
                **inputs,
                max_length=50,
                num_beams=5,  # Better quality with beam search
                early_stopping=True
            )
        
        caption = self.processor.decode(out[0], skip_special_tokens=True)
        return caption
    
    def _generate_git(self, image: Image.Image) -> str:
        """Generate caption using GIT model"""
        # Process image
        inputs = self.processor(images=image, return_tensors="pt").to(self.device)
        
        # Generate caption
        with torch.no_grad():
            generated_ids = self.model.generate(
                pixel_values=inputs.pixel_values,
                max_length=50
            )
        
        caption = self.processor.batch_decode(generated_ids, skip_special_tokens=True)[0]
        return caption
    
    def _clean_alt_text(self, alt_text: str, max_length: int) -> str:
        """Clean and format generated alt text"""
        # Remove common prefixes that BLIP adds
        prefixes_to_remove = [
            "a picture of ",
            "an image of ",
            "a photo of ",
            "there is ",
            "arafed ",  # Common BLIP artifact
        ]
        
        alt_text_lower = alt_text.lower()
        for prefix in prefixes_to_remove:
            if alt_text_lower.startswith(prefix):
                alt_text = alt_text[len(prefix):]
                break
        
        # Capitalize first letter
        if alt_text:
            alt_text = alt_text[0].upper() + alt_text[1:]
        
        # Truncate if needed
        if len(alt_text) > max_length:
            alt_text = alt_text[:max_length-3] + "..."
        
        return alt_text.strip()


class HuggingFaceInferenceAPI:
    """
    Hugging Face Inference API (FREE tier available)
    Falls back to this if local models don't work
    """
    
    def __init__(self, api_token: Optional[str] = None):
        """
        Initialize Hugging Face Inference API
        
        Args:
            api_token: HF token (if None, reads from HF_TOKEN env var)
                      Get free token at: https://huggingface.co/settings/tokens
        """
        self.api_token = api_token or os.getenv("HF_TOKEN")
        self.enabled = False
        
        if not self.api_token:
            print("⚠️  No Hugging Face token found. Set HF_TOKEN environment variable.")
            print("   Get free token at: https://huggingface.co/settings/tokens")
            return
        
        try:
            import requests
            self.requests = requests
            self.enabled = True
            self.api_url = "https://api-inference.huggingface.co/models/Salesforce/blip-image-captioning-base"
            print("✅ Hugging Face Inference API initialized")
        except ImportError:
            print("❌ 'requests' library not available. Run: pip install requests")
    
    def is_enabled(self) -> bool:
        """Check if API is ready"""
        return self.enabled and self.api_token is not None
    
    def generate_alt_text(
        self,
        image_data: bytes,
        shape_name: str = "",
        slide_number: int = 0,
        max_length: int = 250
    ) -> Optional[str]:
        """
        Generate alt text using Hugging Face Inference API
        
        Args:
            image_data: Raw image bytes
            shape_name: Shape name
            slide_number: Slide number
            max_length: Maximum length
            
        Returns:
            Generated alt text or None
        """
        if not self.is_enabled():
            return None
        
        try:
            headers = {"Authorization": f"Bearer {self.api_token}"}
            response = self.requests.post(
                self.api_url,
                headers=headers,
                data=image_data,
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                if isinstance(result, list) and len(result) > 0:
                    caption = result[0].get("generated_text", "")
                    return self._clean_alt_text(caption, max_length)
            else:
                print(f"HF API error: {response.status_code}")
                return None
                
        except Exception as e:
            print(f"HF API request failed: {e}")
            return None
    
    def _clean_alt_text(self, alt_text: str, max_length: int) -> str:
        """Clean generated text"""
        # Remove common prefixes
        prefixes = ["a picture of ", "an image of ", "a photo of "]
        alt_text_lower = alt_text.lower()
        for prefix in prefixes:
            if alt_text_lower.startswith(prefix):
                alt_text = alt_text[len(prefix):]
                break
        
        # Capitalize first letter
        if alt_text:
            alt_text = alt_text[0].upper() + alt_text[1:]
        
        # Truncate if needed
        if len(alt_text) > max_length:
            alt_text = alt_text[:max_length-3] + "..."
        
        return alt_text.strip()


# Singleton instances
_local_model: Optional[LocalVisionModel] = None
_hf_api: Optional[HuggingFaceInferenceAPI] = None


def get_vision_model() -> Optional[LocalVisionModel]:
    """Get or create local vision model singleton"""
    global _local_model
    if _local_model is None:
        model_name = os.getenv("LOCAL_VISION_MODEL", "blip-base")
        _local_model = LocalVisionModel(model_name)
    return _local_model


def get_hf_api() -> Optional[HuggingFaceInferenceAPI]:
    """Get or create Hugging Face API singleton"""
    global _hf_api
    if _hf_api is None:
        _hf_api = HuggingFaceInferenceAPI()
    return _hf_api


def generate_alt_text_free(
    image_data: bytes,
    shape_name: str = "",
    slide_number: int = 0,
    max_length: int = 250
) -> Optional[str]:
    """
    Generate alt text using FREE methods (tries local first, then HF API)
    
    Priority:
    1. Local AI model (completely free, unlimited)
    2. Hugging Face Inference API (free tier)
    3. None (fallback to placeholder in main code)
    
    Args:
        image_data: Raw image bytes
        shape_name: Shape name
        slide_number: Slide number
        max_length: Maximum length
        
    Returns:
        Generated alt text or None
    """
    # Try local model first (best option - free and unlimited)
    local_model = get_vision_model()
    if local_model and local_model.is_enabled():
        result = local_model.generate_alt_text(image_data, shape_name, slide_number, max_length)
        if result:
            return result
    
    # Fallback to Hugging Face API (free tier)
    hf_api = get_hf_api()
    if hf_api and hf_api.is_enabled():
        result = hf_api.generate_alt_text(image_data, shape_name, slide_number, max_length)
        if result:
            return result
    
    # If both fail, return None (main code will use placeholder)
    return None
