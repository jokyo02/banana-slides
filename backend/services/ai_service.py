"""
AI Service - handles all AI model interactions
Based on demo.py and gemini_genai.py
TODO: use structured output API
"""
import os
import json
import re
import requests
from typing import List, Dict, Optional, Union
from textwrap import dedent
from google import genai
from google.genai import types
from PIL import Image
from .prompts import (
    get_outline_generation_prompt,
    get_outline_parsing_prompt,
    get_page_description_prompt,
    get_image_generation_prompt,
    get_image_edit_prompt,
    get_description_to_outline_prompt,
    get_description_split_prompt
)


class AIService:
    """Service for AI model interactions using Gemini"""
    
    def __init__(self, api_key: str, api_base: str = None):
        """Initialize AI service with API credentials"""
        self.client = genai.Client(
            http_options=types.HttpOptions(
                base_url=api_base
            ),
            api_key=api_key
        )
        self.text_model = "gemini-2.5-flash"
        self.image_model = "gemini-3-pro-image-preview"
    
    @staticmethod
    def extract_image_urls_from_markdown(text: str) -> List[str]:
        """
        从 markdown 文本中提取图片 URL
        
        Args:
            text: Markdown 文本，可能包含 ![](url) 格式的图片
            
        Returns:
            图片 URL 列表
        """
        if not text:
            return []
        
        # 匹配 markdown 图片语法: ![](url) 或 ![alt](url)
        pattern = r'!\[.*?\]\((.*?)\)'
        matches = re.findall(pattern, text)
        
        # 过滤掉空字符串和无效 URL
        urls = [url.strip() for url in matches if url.strip() and (url.startswith('http://') or url.startswith('https://'))]
        
        return urls
    
    @staticmethod
    def download_image_from_url(url: str) -> Optional[Image.Image]:
        """
        从 URL 下载图片并返回 PIL Image 对象
        
        Args:
            url: 图片 URL
            
        Returns:
            PIL Image 对象，如果下载失败则返回 None
        """
        try:
            print(f"[DEBUG] Downloading image from URL: {url}")
            response = requests.get(url, timeout=30, stream=True)
            response.raise_for_status()
            
            # 从响应内容创建 PIL Image
            image = Image.open(response.raw)
            # 确保图片被加载
            image.load()
            print(f"[DEBUG] Successfully downloaded image: {image.size}, {image.mode}")
            return image
        except Exception as e:
            print(f"[ERROR] Failed to download image from {url}: {str(e)}")
            return None
    
    def generate_outline(self, idea_prompt: str) -> List[Dict]:
        """
        Generate PPT outline from idea prompt
        Based on demo.py gen_outline()
        
        Returns:
            List of outline items (may contain parts with pages or direct pages)
        """
        outline_prompt = get_outline_generation_prompt(idea_prompt)
        
        response = self.client.models.generate_content(
            model=self.text_model,
            contents=outline_prompt,
            config=types.GenerateContentConfig(
                thinking_config=types.ThinkingConfig(thinking_budget=1000),
            ),
        )
        
        outline_text = response.text.strip().strip("```json").strip("```").strip()
        outline = json.loads(outline_text)
        return outline
    
    def parse_outline_text(self, outline_text: str) -> List[Dict]:
        """
        Parse user-provided outline text into structured outline format
        This method analyzes the text and splits it into pages without modifying the original text
        
        Args:
            outline_text: User-provided outline text (may contain sections, titles, bullet points, etc.)
        
        Returns:
            List of outline items (may contain parts with pages or direct pages)
        """
        parse_prompt = get_outline_parsing_prompt(outline_text)
        
        response = self.client.models.generate_content(
            model=self.text_model,
            contents=parse_prompt,
            config=types.GenerateContentConfig(
                thinking_config=types.ThinkingConfig(thinking_budget=1000),
            ),
        )
        
        outline_json = response.text.strip().strip("```json").strip("```").strip()
        outline = json.loads(outline_json)
        return outline
    
    def flatten_outline(self, outline: List[Dict]) -> List[Dict]:
        """
        Flatten outline structure to page list
        Based on demo.py flatten_outline()
        """
        pages = []
        for item in outline:
            if "part" in item and "pages" in item:
                # This is a part, expand its pages
                for page in item["pages"]:
                    page_with_part = page.copy()
                    page_with_part["part"] = item["part"]
                    pages.append(page_with_part)
            else:
                # This is a direct page
                pages.append(item)
        return pages
    
    def generate_page_description(self, idea_prompt: str, outline: List[Dict], 
                                 page_outline: Dict, page_index: int) -> str:
        """
        Generate description for a single page
        Based on demo.py gen_desc() logic
        
        Args:
            idea_prompt: Original user idea
            outline: Complete outline
            page_outline: Outline for this specific page
            page_index: Page number (1-indexed)
        
        Returns:
            Text description for the page
        """
        part_info = f"\nThis page belongs to: {page_outline['part']}" if 'part' in page_outline else ""
        
        desc_prompt = get_page_description_prompt(
            idea_prompt=idea_prompt,
            outline=outline,
            page_outline=page_outline,
            page_index=page_index,
            part_info=part_info
        )
        
        response = self.client.models.generate_content(
            model=self.text_model,
            contents=desc_prompt,
            config=types.GenerateContentConfig(
                thinking_config=types.ThinkingConfig(thinking_budget=1000),
            ),
        )
        
        page_desc = response.text
        return dedent(page_desc)
    
    def generate_outline_text(self, outline: List[Dict]) -> str:
        """
        Convert outline to text format for prompts
        Based on demo.py gen_outline_text()
        """
        text_parts = []
        for i, item in enumerate(outline, 1):
            if "part" in item and "pages" in item:
                text_parts.append(f"{i}. {item['part']}")
            else:
                text_parts.append(f"{i}. {item.get('title', 'Untitled')}")
        result = "\n".join(text_parts)
        return dedent(result)
    
    def generate_image_prompt(self, outline: List[Dict], page: Dict, 
                            page_desc: str, page_index: int, 
                            has_material_images: bool = False,
                            extra_requirements: Optional[str] = None) -> str:
        """
        Generate image generation prompt for a page
        Based on demo.py gen_prompts()
        
        Args:
            outline: Complete outline
            page: Page outline data
            page_desc: Page description text
            page_index: Page number (1-indexed)
            has_material_images: 是否有素材图片（从项目描述中提取的图片）
            extra_requirements: Optional extra requirements to apply to all pages
        
        Returns:
            Image generation prompt
        """
        outline_text = self.generate_outline_text(outline)
        
        # Determine current section
        if 'part' in page:
            current_section = page['part']
        else:
            current_section = f"{page.get('title', 'Untitled')}"
        
        prompt = get_image_generation_prompt(
            page_desc=page_desc,
            outline_text=outline_text,
            current_section=current_section,
            has_material_images=has_material_images,
            extra_requirements=extra_requirements
        )
        
        return prompt
    
    def generate_image(self, prompt: str, ref_image_path: Optional[str] = None, 
                      aspect_ratio: str = "16:9", resolution: str = "2K",
                      additional_ref_images: Optional[List[Union[str, Image.Image]]] = None) -> Optional[Image.Image]:
        """
        Generate image using Gemini image model
        Based on gemini_genai.py gen_image()
        
        Args:
            prompt: Image generation prompt
            ref_image_path: Path to reference image (optional). If None, will generate based on prompt only.
            aspect_ratio: Image aspect ratio (currently not used, kept for compatibility)
            resolution: Image resolution (currently not used, kept for compatibility)
            additional_ref_images: 额外的参考图片列表，可以是本地路径、URL 或 PIL Image 对象
        
        Returns:
            PIL Image object or None if failed
        
        Raises:
            Exception with detailed error message if generation fails
        """
        try:
            print(f"[DEBUG] Generating image with prompt (first 100 chars): {prompt[:1000]}...")
            print(f"[DEBUG] Reference image: {ref_image_path}")
            if additional_ref_images:
                print(f"[DEBUG] Additional reference images: {len(additional_ref_images)}")
            print(f"[DEBUG] Config - aspect_ratio: {aspect_ratio}, resolution: {resolution}")

            # 构建 contents 列表，包含 prompt 和所有参考图片
            # 约定：如果有主参考图，则放在第一个索引，其后是文本 prompt，再后是其他参考图
            contents = []
            
            # 添加主参考图片（如果提供了路径，放在第一个位置）
            if ref_image_path:
                if not os.path.exists(ref_image_path):
                    raise FileNotFoundError(f"Reference image not found: {ref_image_path}")
                main_ref_image = Image.open(ref_image_path)
                contents.append(main_ref_image)
            
            # 文本 prompt 紧跟在主参考图之后（或成为第一个元素）
            contents.append(prompt)
            
            # 添加额外的参考图片
            if additional_ref_images:
                for ref_img in additional_ref_images:
                    if isinstance(ref_img, Image.Image):
                        # 已经是 PIL Image 对象
                        contents.append(ref_img)
                    elif isinstance(ref_img, str):
                        # 可能是本地路径或 URL
                        if os.path.exists(ref_img):
                            # 本地路径
                            contents.append(Image.open(ref_img))
                        elif ref_img.startswith('http://') or ref_img.startswith('https://'):
                            # URL，需要下载
                            downloaded_img = self.download_image_from_url(ref_img)
                            if downloaded_img:
                                contents.append(downloaded_img)
                            else:
                                print(f"[WARNING] Failed to download image from URL: {ref_img}, skipping...")
                        else:
                            print(f"[WARNING] Invalid image reference: {ref_img}, skipping...")
            
            print(f"[DEBUG] Calling Gemini API for image generation with {len(contents) - 1} reference images...")
            response = self.client.models.generate_content(
                model=self.image_model,
                contents=contents,
                config=types.GenerateContentConfig(
                    response_modalities=['TEXT', 'IMAGE'],
                    image_config=types.ImageConfig(
                        aspect_ratio=aspect_ratio,
                        image_size=resolution
                    ),
                )
            )
            print(f"[DEBUG] Gemini API call completed")
            
            print(f"[DEBUG] API response received, checking parts...")
            for i, part in enumerate(response.parts):
                if part.text is not None:   
                    print(f"[DEBUG] Part {i}: TEXT - {part.text[:100]}")
                else:
                    # Try to get image from part
                    try:
                        print(f"[DEBUG] Part {i}: Attempting to extract image...")
                        image = part.as_image()
                        if image:
                            # Don't check image.size - it might not be a standard PIL Image yet
                            print(f"[DEBUG] Successfully extracted image from part {i}")
                            return image
                    except Exception as e:
                        print(f"[DEBUG] Part {i}: Failed to extract image - {str(e)}")
            
            # If we get here, no image was found in the response
            error_msg = "No image found in API response. "
            if response.parts:
                error_msg += f"Response had {len(response.parts)} parts but none contained valid images."
            else:
                error_msg += "Response had no parts."
            
            raise ValueError(error_msg)
            
        except Exception as e:
            error_detail = f"Error generating image: {type(e).__name__}: {str(e)}"
            print(f"[ERROR] {error_detail}")
            import traceback
            traceback.print_exc()
            raise Exception(error_detail) from e
    
    def edit_image(self, prompt: str, current_image_path: str,
                  aspect_ratio: str = "16:9", resolution: str = "2K",
                  original_description: str = None,
                  additional_ref_images: Optional[List[Union[str, Image.Image]]] = None) -> Optional[Image.Image]:
        """
        Edit existing image with natural language instruction
        Uses current image as reference
        
        Args:
            prompt: Edit instruction
            current_image_path: Path to current page image
            aspect_ratio: Image aspect ratio
            resolution: Image resolution
            original_description: Original page description to include in prompt
            additional_ref_images: 额外的参考图片列表，可以是本地路径、URL 或 PIL Image 对象
        
        Returns:
            PIL Image object or None if failed
        """
        # Build edit instruction with original description if available
        edit_instruction = get_image_edit_prompt(
            edit_instruction=prompt,
            original_description=original_description
        )
        return self.generate_image(edit_instruction, current_image_path, aspect_ratio, resolution, additional_ref_images)
    
    def parse_description_to_outline(self, description_text: str) -> List[Dict]:
        """
        从描述文本解析出大纲结构
        
        Args:
            description_text: 用户提供的完整页面描述文本
        
        Returns:
            List of outline items (may contain parts with pages or direct pages)
        """
        parse_prompt = get_description_to_outline_prompt(description_text)
        
        response = self.client.models.generate_content(
            model=self.text_model,
            contents=parse_prompt,
            config=types.GenerateContentConfig(
                thinking_config=types.ThinkingConfig(thinking_budget=1000),
            ),
        )
        
        outline_json = response.text.strip().strip("```json").strip("```").strip()
        outline = json.loads(outline_json)
        return outline
    
    def parse_description_to_page_descriptions(self, description_text: str, outline: List[Dict]) -> List[str]:
        """
        从描述文本切分出每页描述
        
        Args:
            description_text: 用户提供的完整页面描述文本
            outline: 已解析出的大纲结构
        
        Returns:
            List of page descriptions (strings), one for each page in the outline
        """
        split_prompt = get_description_split_prompt(description_text, outline)
        
        response = self.client.models.generate_content(
            model=self.text_model,
            contents=split_prompt,
            config=types.GenerateContentConfig(
                thinking_config=types.ThinkingConfig(thinking_budget=1000),
            ),
        )
        
        descriptions_json = response.text.strip().strip("```json").strip("```").strip()
        descriptions = json.loads(descriptions_json)
        
        # 确保返回的是字符串列表
        if isinstance(descriptions, list):
            return [str(desc) for desc in descriptions]
        else:
            raise ValueError("Expected a list of page descriptions, but got: " + str(type(descriptions)))

