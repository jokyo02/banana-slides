import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Sparkles, FileText, FileEdit } from 'lucide-react';
import { Button, Input, Textarea, Card, useToast } from '@/components/shared';
import { useProjectStore } from '@/store/useProjectStore';
import * as api from '@/api/endpoints';
import { getImageUrl } from '@/api/client';
import type { UserTemplate } from '@/api/endpoints';

type CreationType = 'idea' | 'outline' | 'description';

const templates = [
  { id: '1', name: '简约商务', preview: '' },
  { id: '2', name: '活力色彩', preview: '' },
  { id: '3', name: '科技蓝', preview: '' },
];

export const Home: React.FC = () => {
  const navigate = useNavigate();
  const { initializeProject, isGlobalLoading } = useProjectStore();
  const { show, ToastContainer } = useToast();
  
  const [activeTab, setActiveTab] = useState<CreationType>('idea');
  const [content, setContent] = useState('');
  const [selectedTemplate, setSelectedTemplate] = useState<File | null>(null);
  const [selectedTemplateId, setSelectedTemplateId] = useState<string | null>(null);
  const [userTemplates, setUserTemplates] = useState<UserTemplate[]>([]);
  const [isLoadingTemplates, setIsLoadingTemplates] = useState(false);

  // 加载用户模板列表
  useEffect(() => {
    loadUserTemplates();
  }, []);

  const tabConfig = {
    idea: {
      icon: <Sparkles size={20} />,
      label: '一句话生成',
      placeholder: '例如：生成一份关于 AI 发展史的演讲 PPT',
      description: '输入你的想法，AI 将为你生成完整的 PPT',
    },
    outline: {
      icon: <FileText size={20} />,
      label: '从大纲生成（未完善，但可手动编辑大纲卡片）',
      placeholder: '粘贴你的 PPT 大纲...\n\n例如：\n第一部分：AI 的起源\n- 1950 年代的开端\n- 达特茅斯会议\n\n第二部分：发展历程\n...',
      description: '已有大纲？直接粘贴即可快速生成',
    },
    description: {
      icon: <FileEdit size={20} />,
      label: '从描述生成（未完善，但可手动编辑描述卡片）',
      placeholder: '粘贴你的详细页面描述...\n\n例如：\n第 1 页\n标题：人工智能的诞生\n内容：1950 年，图灵提出"图灵测试"...\n...',
      description: '已有完整描述？直接生成图片',
    },
  };

  const loadUserTemplates = async () => {
    setIsLoadingTemplates(true);
    try {
      const response = await api.listUserTemplates();
      if (response.data?.templates) {
        setUserTemplates(response.data.templates);
      }
    } catch (error: any) {
      console.error('加载用户模板失败:', error);
      // 静默失败，不影响主流程
    } finally {
      setIsLoadingTemplates(false);
    }
  };

  const handleTemplateUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      try {
        // 上传到后端
        const response = await api.uploadUserTemplate(file);
        if (response.data) {
          const template = response.data;
          setUserTemplates(prev => [template, ...prev]);
          setSelectedTemplateId(template.template_id);
          setSelectedTemplate(null); // 清空本地文件选择
          show({ message: '模板上传成功', type: 'success' });
        }
      } catch (error: any) {
        console.error('上传模板失败:', error);
        show({ message: '模板上传失败: ' + (error.message || '未知错误'), type: 'error' });
      }
    }
  };

  const handleSelectUserTemplate = (template: UserTemplate) => {
    setSelectedTemplateId(template.template_id);
    setSelectedTemplate(null); // 清空本地文件选择
  };

  const handleRemoveTemplate = () => {
    setSelectedTemplateId(null);
    setSelectedTemplate(null);
  };

  const handleSubmit = async () => {
    if (!content.trim()) {
      show({ message: '请输入内容', type: 'error' });
      return;
    }

    try {
      // 如果有选中的用户模板，需要先获取模板文件
      let templateFile: File | undefined = selectedTemplate || undefined;
      
      if (selectedTemplateId && !templateFile) {
        // 从用户模板创建 File 对象
        // 注意：这里我们需要从 URL 获取图片并转换为 File
        // 但为了简化，我们可以先使用 selectedTemplateId 来标识
        // 实际上，在创建项目时，我们可以通过模板 ID 来关联
        // 但目前的 initializeProject 只接受 File，所以我们需要获取模板图片
        try {
          const template = userTemplates.find(t => t.template_id === selectedTemplateId);
          if (template) {
            const imageUrl = getImageUrl(template.template_image_url);
            const response = await fetch(imageUrl);
            const blob = await response.blob();
            templateFile = new File([blob], 'template.png', { type: blob.type });
          }
        } catch (error) {
          console.warn('获取模板文件失败:', error);
        }
      }

      await initializeProject(activeTab, content, templateFile);
      
      // 根据类型跳转到不同页面
      const projectId = localStorage.getItem('currentProjectId');
      if (!projectId) {
        show({ message: '项目创建失败', type: 'error' });
        return;
      }
      
      if (activeTab === 'idea' || activeTab === 'outline') {
        navigate(`/project/${projectId}/outline`);
      } else {
        navigate(`/project/${projectId}/detail`);
      }
    } catch (error: any) {
      console.error('创建项目失败:', error);
      // 错误已经在 store 中处理并显示
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-banana-50 via-white to-gray-50">
      {/* 导航栏 */}
      <nav className="h-16 bg-white shadow-sm border-b border-gray-100">
        <div className="max-w-7xl mx-auto px-4 h-full flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-10 h-10 bg-gradient-to-br from-banana-500 to-banana-600 rounded-lg flex items-center justify-center text-2xl">
              🍌
            </div>
            <span className="text-xl font-bold text-gray-900">蕉幻</span>
          </div>
          <div className="flex items-center gap-4">
            <Button variant="ghost" size="sm" onClick={() => navigate('/history')}>
              历史项目
            </Button>
            <Button variant="ghost" size="sm">帮助</Button>
          </div>
        </div>
      </nav>

      {/* 主内容 */}
      <main className="max-w-4xl mx-auto px-4 py-16">
        {/* 标题区 */}
        <div className="text-center mb-12">
          <h1 className="text-5xl font-bold text-gray-900 mb-4">
            🍌 蕉幻 Banana Slides
          </h1>
          <p className="text-xl text-gray-600">
            AI 原生 PPT 生成器，一句话创造精彩
          </p>
        </div>

        {/* 创建卡片 */}
        <Card className="p-10">
          {/* 选项卡 */}
          <div className="flex gap-4 mb-8">
            {(Object.keys(tabConfig) as CreationType[]).map((type) => {
              const config = tabConfig[type];
              return (
                <button
                  key={type}
                  onClick={() => setActiveTab(type)}
                  className={`flex-1 flex items-center justify-center gap-2 px-6 py-3 rounded-lg font-medium transition-all ${
                    activeTab === type
                      ? 'bg-gradient-to-r from-banana-500 to-banana-600 text-black shadow-yellow'
                      : 'bg-white border border-gray-200 text-gray-700 hover:bg-banana-50'
                  }`}
                >
                  {config.icon}
                  {config.label}
                </button>
              );
            })}
          </div>

          {/* 描述 */}
          <p className="text-gray-600 mb-6">
            {tabConfig[activeTab].description}
          </p>

          {/* 输入区 */}
          <Textarea
            placeholder={tabConfig[activeTab].placeholder}
            value={content}
            onChange={(e) => setContent(e.target.value)}
            rows={activeTab === 'idea' ? 4 : 10}
            className="mb-6"
          />

          {/* 模板选择 */}
          <div className="mb-8">
            <h3 className="text-lg font-semibold text-gray-900 mb-4">
              🎨 选择风格模板 (可选)
            </h3>
            
            {/* 用户已保存的模板 */}
            {userTemplates.length > 0 && (
              <div className="mb-4">
                <h4 className="text-sm font-medium text-gray-700 mb-2">我的模板</h4>
                <div className="grid grid-cols-4 gap-4 mb-4">
                  {userTemplates.map((template) => (
                    <div
                      key={template.template_id}
                      onClick={() => handleSelectUserTemplate(template)}
                      className={`aspect-[4/3] rounded-lg border-2 cursor-pointer transition-all relative overflow-hidden ${
                        selectedTemplateId === template.template_id
                          ? 'border-banana-500 ring-2 ring-banana-200'
                          : 'border-gray-200 hover:border-banana-300'
                      }`}
                    >
                      <img
                        src={getImageUrl(template.template_image_url)}
                        alt={template.name || 'Template'}
                        className="absolute inset-0 w-full h-full object-cover"
                      />
                      {selectedTemplateId === template.template_id && (
                        <div className="absolute inset-0 bg-banana-500 bg-opacity-20 flex items-center justify-center">
                          <span className="text-white font-semibold text-sm">已选择</span>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="grid grid-cols-4 gap-4">
              {/* 预设模板 */}
              {templates.map((template) => (
                <div
                  key={template.id}
                  className="aspect-[4/3] rounded-lg border-2 border-gray-200 hover:border-banana-500 cursor-pointer transition-all bg-gray-100 flex items-center justify-center"
                >
                  <span className="text-sm text-gray-500">{template.name}</span>
                </div>
              ))}

              {/* 上传新模板 */}
              <label className="aspect-[4/3] rounded-lg border-2 border-dashed border-gray-300 hover:border-banana-500 cursor-pointer transition-all flex flex-col items-center justify-center gap-2 relative overflow-hidden">
                <span className="text-2xl">+</span>
                <span className="text-sm text-gray-500">上传模板</span>
                <input
                  type="file"
                  accept="image/*"
                  onChange={handleTemplateUpload}
                  className="hidden"
                  disabled={isLoadingTemplates}
                />
              </label>
            </div>

            {/* 显示已选择的模板提示 */}
            {selectedTemplateId && (
              <div className="mt-4 flex items-center justify-between p-3 bg-banana-50 rounded-lg">
                <span className="text-sm text-gray-700">
                  已选择模板: {userTemplates.find(t => t.template_id === selectedTemplateId)?.name || '未命名模板'}
                </span>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={handleRemoveTemplate}
                  className="text-gray-600 hover:text-gray-900"
                >
                  取消选择
                </Button>
              </div>
            )}
          </div>

          {/* 提交按钮 */}
          <div className="flex justify-center">
            <Button
              size="lg"
              onClick={handleSubmit}
              loading={isGlobalLoading}
              className="w-64"
            >
              开始生成
            </Button>
          </div>
        </Card>
      </main>
      <ToastContainer />
    </div>
  );
};

