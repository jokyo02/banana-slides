"""
Agent Controller - handles agent-related endpoints for slide editing
"""
import os
import logging
from google.genai import types
from flask import Blueprint, request, current_app
from models import db, Project, Page
from utils import success_response, error_response, not_found, bad_request
from agno.agent import Agent
from agno.models.google import Gemini
from agno.tools.sleep import SleepTools
from services.slide_agent_tools import SlideAgentTools

logger = logging.getLogger(__name__)

agent_bp = Blueprint('agent', __name__, url_prefix='/api/projects')


gemini_model = Gemini(
    id=os.getenv("AGENT_MODEL", "gemini-2.5-flash"),
    client_params={
        "http_options": types.HttpOptions(
            base_url=os.getenv("GOOGLE_API_BASE")
        ),
        "api_key": os.getenv("GOOGLE_API_KEY", "")
    }
)

def create_slide_agent(project_id: str, app=None) -> Agent:
    """Create an Agno agent for slide editing"""
    
    with app.app_context():
        api_key = current_app.config.get('GOOGLE_API_KEY', '')
    
    if not api_key:
        raise ValueError("API key not configured. Please set GOOGLE_API_KEY environment variable")
    
    # Create tools
    slide_tools = SlideAgentTools(project_id=project_id, app=app)
    sleep_tools = SleepTools()
    
    # Get project context（复用 SlideAgentTools 的逻辑）
    with app.app_context():
        project_pages_result = slide_tools.get_project_pages()
        if not project_pages_result.get("success"):
            raise ValueError(project_pages_result.get("message", "Failed to load project pages"))
        
        pages = project_pages_result.get("pages", [])
        pages_info = []
        for p in pages:
            pages_info.append({
                "index": p.get("index"),
                "page_id": p.get("page_id"),
                "title": p.get("title"),
                "has_description": bool(p.get("has_description")),
                "has_image": bool(p.get("has_image")),
                "status": p.get("status"),
            })
    
    instructions = f"""你是一个智能幻灯片编辑助手，可以帮助用户编辑和管理「整个项目」的所有幻灯片，而不仅仅是当前预览的页面。

## 你的能力

你拥有以下工具（可对任意页面使用，不限于当前页）：
1. **get_project_pages** - 获取整个项目的所有页面信息（标题、大纲、描述、图片状态等），用于全局规划和批量修改
2. **edit_page_image** - 编辑指定页面图片（使用自然语言指令）
3. **update_page_description** - 更新指定页面的描述内容
4. **update_page_outline** - 更新指定页面的大纲内容
5. **regenerate_page_image** - 重新生成指定页面的图片
6. **sleep** - 等待一段时间（当任务队列已满时使用）

## 重要规则

1. **并发任务限制**：一次最多只能执行4个异步任务（edit_page_image 和 regenerate_page_image 是异步的，可以作用于多个页面）
2. **任务队列管理**：如果当前有4个任务在执行，你需要使用 sleep 工具等待一段时间，然后再继续对其他页面发起任务
3. **异步任务**：edit_page_image 和 regenerate_page_image 会返回 task_id，任务在后台执行，你可以并行处理多个页面
4. **同步任务**：update_page_description 和 update_page_outline 是同步的，会立即完成，适合对多个页面依次批量更新

## 当前项目信息（仅作为初始参考，实际操作时可以随时调用 get_project_pages 获取最新数据）

项目ID: {project_id}
总页数: {len(pages_info)}

页面列表:
{chr(10).join([f"  {p['index']}. {p['title']} (ID: {p['page_id']}) - 描述: {'有' if p['has_description'] else '无'}, 图片: {'有' if p['has_image'] else '无'}, 状态: {p['status']}" for p in pages_info])}

## 使用建议

- 当用户提出「整体优化」或「全局修改」需求时：
  - 先调用 **get_project_pages** 获取所有页面的信息和结构
  - 根据用户意图，选择一组页面（可以按索引或标题识别）进行批量修改
- 当用户要求编辑某些页面的图片时，使用 **edit_page_image**，必要时对多页重复调用（注意并发限制和 sleep）
- 当用户要求统一调整多页描述时，使用 **update_page_description**，逐页更新
- 当用户要求统一调整多页大纲结构时，使用 **update_page_outline**，逐页更新
- 当用户要求重新生成多页图片时，使用 **regenerate_page_image**，可以对多个页面发起任务
- 如果任务队列已满，使用 **sleep** 等待后再继续对剩余页面发起任务
"""
    
    agent = Agent(
        name="SlideEditor",
        model=gemini_model,
        tools=[slide_tools, sleep_tools],
        instructions=instructions,
        markdown=True,
        debug_mode=False
    )
    
    return agent


@agent_bp.route('/<project_id>/agent/chat', methods=['POST'])
def agent_chat(project_id):
    """
    POST /api/projects/{project_id}/agent/chat - Chat with slide editing agent
    
    Request body:
    {
        "message": "请编辑第1页的图片，将背景改为蓝色"
    }
    """
    try:
        project = Project.query.get(project_id)
        if not project:
            return not_found('Project')
        
        data = request.get_json()
        if not data or 'message' not in data:
            return bad_request("message is required")
        
        message = data['message']
        
        # Get app instance for agent creation
        app_instance = current_app._get_current_object()
        
        # Create agent with app context
        agent = create_slide_agent(project_id, app=app_instance)
        
        # Get response from agent
        # Use run method which returns a RunResponse
        # Run within app context to ensure database access works
        with app_instance.app_context():
            run_response = agent.run(message)
        
        # Extract content from response
        # The response object should have a content attribute
        response_text = ""
        if hasattr(run_response, 'content'):
            response_text = run_response.content
        elif hasattr(run_response, 'messages') and getattr(run_response, "messages", None):
            # Get last assistant message with string content
            for msg in reversed(run_response.messages):
                content = getattr(msg, "content", None)
                if isinstance(content, str) and content.strip():
                    response_text = content
                    break
        else:
            response_text = str(run_response)
        
        # 注意：run_response.messages 里是复杂的 Message 对象，不能直接 JSON 序列化
        # 前端目前只用到 response 文本，这里先只返回字符串，避免报错
        return success_response({
            "response": response_text
        })
    
    except Exception as e:
        logger.error(f"Agent chat error: {e}", exc_info=True)
        return error_response('AGENT_ERROR', str(e), 500)

