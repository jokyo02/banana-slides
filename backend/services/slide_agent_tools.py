"""
Slide Agent Tools for Agno Framework
Tools for editing slides, updating descriptions, outlines, and regenerating pages.
"""
import os
import logging
import uuid
from typing import Optional, Dict, Any, List
from agno.tools import Toolkit
from flask import current_app
from models import db, Project, Page, Task
from services import AIService, FileService
from services.task_manager import task_manager, generate_single_page_image_task, edit_page_image_task
from datetime import datetime

logger = logging.getLogger(__name__)


class SlideAgentTools(Toolkit):
    """Tools for slide editing agent"""
    
    def __init__(self, project_id: str, app=None):
        # 显式向 Agno 注册要暴露给大模型的工具方法
        # 注意：tools 参数里的可调用对象名（. __name__）会作为工具名称，
        # 需要与我们在 instructions 中提到的名字一致：
        # - edit_page_image
        # - update_page_description
        # - update_page_outline
        # - regenerate_page_image
        # - get_project_pages
        super().__init__(
            name="slide_agent_tools",
            tools=[
                self.edit_page_image,
                self.update_page_description,
                self.update_page_outline,
                self.regenerate_page_image,
                self.get_project_pages,
            ],
        )
        self.project_id = project_id
        self.app = app
        self.active_tasks = []  # Track active async tasks (max 4)
        self.max_concurrent_tasks = 4
    
    def _check_task_limit(self) -> bool:
        """Check if we can start a new task"""
        # Remove completed tasks by checking their status in database
        with self.app.app_context():
            active_count = 0
            for task_info in self.active_tasks:
                task = Task.query.get(task_info.get('task_id'))
                if task and task.status in ['PENDING', 'PROCESSING']:
                    active_count += 1
                else:
                    # Mark as completed
                    task_info['status'] = 'completed'
            
            # Clean up completed tasks
            self.active_tasks = [t for t in self.active_tasks if t.get('status') != 'completed']
            return len([t for t in self.active_tasks if t.get('status') != 'completed']) < self.max_concurrent_tasks
    
    def _add_task(self, task_id: str, task_type: str, page_id: str):
        """Add a task to active tasks list"""
        self.active_tasks.append({
            'task_id': task_id,
            'type': task_type,
            'page_id': page_id,
            'status': 'pending'
        })
    
    def edit_page_image(
        self,
        page_id: str,
        edit_instruction: str,
        use_template: bool = False,
        desc_image_urls: List[str] = None
    ) -> Dict[str, Any]:
        """
        Edit a page image using natural language instructions.
        
        Args:
            page_id: The page ID to edit
            edit_instruction: Natural language instruction for editing
            use_template: Whether to use template image as context
            desc_image_urls: List of description image URLs to use as context
        
        Returns:
            Dict with task_id if async, or success message
        """
        try:
            if not self._check_task_limit():
                return {
                    "success": False,
                    "message": "已达到最大并发任务数(4个)，请等待部分任务完成后再试",
                    "suggestion": "使用sleep工具等待一段时间"
                }
            
            with self.app.app_context():
                page = Page.query.get(page_id)
                if not page or page.project_id != self.project_id:
                    return {"success": False, "message": "页面不存在"}
                
                if not page.generated_image_path:
                    return {"success": False, "message": "页面还没有生成图片，请先生成图片"}
                
                project = Project.query.get(self.project_id)
                if not project:
                    return {"success": False, "message": "项目不存在"}
                
                # Initialize services
                ai_service = AIService(
                    current_app.config['GOOGLE_API_KEY'],
                    current_app.config['GOOGLE_API_BASE']
                )
                file_service = FileService(current_app.config['UPLOAD_FOLDER'])
                
                # Get original description
                original_description = None
                desc_content = page.get_description_content()
                if desc_content:
                    original_description = desc_content.get('text') or ''
                    if not original_description and desc_content.get('text_content'):
                        if isinstance(desc_content['text_content'], list):
                            original_description = '\n'.join(desc_content['text_content'])
                        else:
                            original_description = str(desc_content['text_content'])
                
                # Collect additional reference images
                additional_ref_images = []
                if use_template:
                    template_path = file_service.get_template_path(self.project_id)
                    if template_path:
                        additional_ref_images.append(template_path)
                
                if desc_image_urls:
                    additional_ref_images.extend(desc_image_urls)
                
                # Create task
                task_id = str(uuid.uuid4())
                task = Task(
                    id=task_id,
                    project_id=self.project_id,
                    task_type='EDIT_PAGE_IMAGE',
                    status='PENDING',
                    created_at=datetime.utcnow()
                )
                db.session.add(task)
                db.session.commit()
                
                # Submit background task
                task_manager.submit_task(
                    task_id,
                    edit_page_image_task,
                    self.project_id,
                    page_id,
                    edit_instruction,
                    ai_service,
                    file_service,
                    aspect_ratio=current_app.config.get('DEFAULT_ASPECT_RATIO', '16:9'),
                    resolution=current_app.config.get('DEFAULT_RESOLUTION', '2K'),
                    original_description=original_description,
                    additional_ref_images=additional_ref_images if additional_ref_images else None,
                    temp_dir=None,
                    app=self.app
                )
                
                self._add_task(task_id, 'edit_image', page_id)
                
                return {
                    "success": True,
                    "message": f"已开始编辑页面图片，任务ID: {task_id}",
                    "task_id": task_id
                }
                    
        except Exception as e:
            logger.error(f"Error editing page image: {e}", exc_info=True)
            return {
                "success": False,
                "message": f"编辑失败: {str(e)}"
            }
    
    def update_page_description(
        self,
        page_id: str,
        description_content: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Update page description content.
        
        Args:
            page_id: The page ID to update
            description_content: Description content dict with title, text_content, layout_suggestion
        
        Returns:
            Dict with success status and message
        """
        try:
            with self.app.app_context():
                page = Page.query.get(page_id)
                if not page or page.project_id != self.project_id:
                    return {"success": False, "message": "页面不存在"}
                
                page.set_description_content(description_content)
                page.updated_at = datetime.utcnow()
                
                project = Project.query.get(self.project_id)
                if project:
                    project.updated_at = datetime.utcnow()
                
                db.session.commit()
                
                return {
                    "success": True,
                    "message": "页面描述已更新"
                }
                    
        except Exception as e:
            logger.error(f"Error updating page description: {e}", exc_info=True)
            return {
                "success": False,
                "message": f"更新失败: {str(e)}"
            }
    
    def update_page_outline(
        self,
        page_id: str,
        outline_content: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Update page outline content.
        
        Args:
            page_id: The page ID to update
            outline_content: Outline content dict with title, points, etc.
        
        Returns:
            Dict with success status and message
        """
        try:
            with self.app.app_context():
                page = Page.query.get(page_id)
                if not page or page.project_id != self.project_id:
                    return {"success": False, "message": "页面不存在"}
                
                page.set_outline_content(outline_content)
                page.updated_at = datetime.utcnow()
                
                project = Project.query.get(self.project_id)
                if project:
                    project.updated_at = datetime.utcnow()
                
                db.session.commit()
                
                return {
                    "success": True,
                    "message": "页面大纲已更新"
                }
                    
        except Exception as e:
            logger.error(f"Error updating page outline: {e}", exc_info=True)
            return {
                "success": False,
                "message": f"更新失败: {str(e)}"
            }
    
    def regenerate_page_image(
        self,
        page_id: str,
        force_regenerate: bool = True
    ) -> Dict[str, Any]:
        """
        Regenerate page image.
        
        Args:
            page_id: The page ID to regenerate
            force_regenerate: Whether to force regenerate even if image exists
        
        Returns:
            Dict with task_id if async, or success message
        """
        try:
            if not self._check_task_limit():
                return {
                    "success": False,
                    "message": "已达到最大并发任务数(4个)，请等待部分任务完成后再试",
                    "suggestion": "使用sleep工具等待一段时间"
                }
            
            with self.app.app_context():
                page = Page.query.get(page_id)
                if not page or page.project_id != self.project_id:
                    return {"success": False, "message": "页面不存在"}
                
                desc_content = page.get_description_content()
                if not desc_content:
                    return {"success": False, "message": "页面还没有描述内容，请先生成描述"}
                
                project = Project.query.get(self.project_id)
                if not project:
                    return {"success": False, "message": "项目不存在"}
                
                # Initialize services
                ai_service = AIService(
                    current_app.config['GOOGLE_API_KEY'],
                    current_app.config['GOOGLE_API_BASE']
                )
                file_service = FileService(current_app.config['UPLOAD_FOLDER'])
                
                # Reconstruct full outline
                all_pages = Page.query.filter_by(project_id=self.project_id).order_by(Page.order_index).all()
                outline = []
                for p in all_pages:
                    oc = p.get_outline_content()
                    if oc:
                        page_data = oc.copy()
                        if p.part:
                            page_data['part'] = p.part
                        outline.append(page_data)
                
                # Create task
                task_id = str(uuid.uuid4())
                task = Task(
                    id=task_id,
                    project_id=self.project_id,
                    task_type='GENERATE_PAGE_IMAGE',
                    status='PENDING',
                    created_at=datetime.utcnow()
                )
                db.session.add(task)
                db.session.commit()
                
                # Submit background task
                task_manager.submit_task(
                    task_id,
                    generate_single_page_image_task,
                    self.project_id,
                    page_id,
                    ai_service,
                    file_service,
                    outline,
                    use_template=True,
                    aspect_ratio=current_app.config.get('DEFAULT_ASPECT_RATIO', '16:9'),
                    resolution=current_app.config.get('DEFAULT_RESOLUTION', '2K'),
                    app=self.app,
                    extra_requirements=project.extra_requirements
                )
                
                self._add_task(task_id, 'regenerate_image', page_id)
                
                return {
                    "success": True,
                    "message": f"已开始重新生成页面图片，任务ID: {task_id}",
                    "task_id": task_id
                }
                    
        except Exception as e:
            logger.error(f"Error regenerating page image: {e}", exc_info=True)
            return {
                "success": False,
                "message": f"生成失败: {str(e)}"
            }

    def get_project_pages(self) -> Dict[str, Any]:
        """
        获取整个项目的页面列表和关键信息，用于全局性修改和规划。
        
        Returns:
            {
              "success": True,
              "project": {...},
              "pages": [
                {
                  "index": 1,
                  "page_id": "...",
                  "title": "...",
                  "outline": {...} | null,
                  "description": {...} | null,
                  "has_image": bool,
                  "status": "..."
                },
                ...
              ]
            }
        """
        try:
            with self.app.app_context():
                project = Project.query.get(self.project_id)
                if not project:
                    return {"success": False, "message": "项目不存在"}

                # 复用已有的模型序列化逻辑，避免到处手写 ORM 查询和字段拼装
                project_dict = project.to_dict(include_pages=True)
                raw_pages = project_dict.get("pages", []) or []

                pages_info: List[Dict[str, Any]] = []
                for i, page in enumerate(raw_pages, 1):
                    outline = page.get("outline_content")
                    desc = page.get("description_content")
                    pages_info.append({
                        "index": i,
                        "page_id": page.get("page_id"),
                        "title": (outline or {}).get("title") or f"页面 {i}",
                        "outline": outline,
                        "description": desc,
                        "has_description": bool(desc),
                        "has_image": bool(page.get("generated_image_url")),
                        "status": page.get("status"),
                    })

                return {
                    "success": True,
                    "project": {
                        "project_id": project_dict.get("project_id"),
                        "title": project_dict.get("idea_prompt"),  # 暂无正式标题字段，先用 idea_prompt 代表
                        "status": project_dict.get("status"),
                        "pages_count": len(pages_info),
                    },
                    "pages": pages_info,
                }
        except Exception as e:
            logger.error(f"Error getting project pages: {e}", exc_info=True)
            return {
                "success": False,
                "message": f"获取项目页面信息失败: {str(e)}"
            }

