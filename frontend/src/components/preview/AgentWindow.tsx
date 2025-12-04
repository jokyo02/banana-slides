import React, { useState, useRef, useEffect } from 'react';
import { MessageCircle, X, Send, Loader2 } from 'lucide-react';
import { Button, useToast } from '@/components/shared';
import { agentChat } from '@/api/endpoints';
import { useParams } from 'react-router-dom';

interface Message {
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
}

interface AgentWindowProps {
  isOpen: boolean;
  onClose: () => void;
}

export const AgentWindow: React.FC<AgentWindowProps> = ({ isOpen, onClose }) => {
  const { projectId } = useParams<{ projectId: string }>();
  const [messages, setMessages] = useState<Message[]>([]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const { show } = useToast();
  const [isEntering, setIsEntering] = useState(false);

  // 滚动到底部
  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    if (isOpen) {
      // 进入时启动缓冲动画
      setIsEntering(false);
      // 下一帧再切换到进入状态，触发过渡
      const timer = requestAnimationFrame(() => {
        setIsEntering(true);
        scrollToBottom();
      });
      return () => cancelAnimationFrame(timer);
    } else {
      setIsEntering(false);
    }
  }, [messages, isOpen]);

  const handleSend = async () => {
    if (!input.trim() || !projectId || isLoading) return;

    const userMessage: Message = {
      role: 'user',
      content: input.trim(),
      timestamp: new Date(),
    };

    setMessages((prev) => [...prev, userMessage]);
    setInput('');
    setIsLoading(true);

    try {
      const response = await agentChat(projectId, userMessage.content);
      
      if (response.success && response.data) {
        const assistantMessage: Message = {
          role: 'assistant',
          content: response.data.response || '收到',
          timestamp: new Date(),
        };
        setMessages((prev) => [...prev, assistantMessage]);
      } else {
        show({
          message: response.error || 'Agent响应失败',
          type: 'error',
        });
      }
    } catch (err: any) {
      console.error('Agent chat error:', err);
      // 兼容 AxiosError / Error / string 等多种情况
      const errorObj = err as any;
      let apiMessage = '发送失败';
      if (errorObj?.response?.data?.error?.message) {
        apiMessage = errorObj.response.data.error.message;
      } else if (errorObj?.response?.data?.message) {
        apiMessage = errorObj.response.data.message;
      } else if (typeof errorObj?.message === 'string') {
        apiMessage = errorObj.message;
      }
      show({
        message: apiMessage,
        type: 'error',
      });
    } finally {
      setIsLoading(false);
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  if (!isOpen) return null;

  return (
    <div
      className={`h-full w-full flex flex-col bg-gradient-to-b from-white via-gray-50 to-gray-100 border-l border-gray-200 shadow-xl transform transition-transform transition-opacity duration-200 ease-out ${
        isEntering ? 'translate-x-0 opacity-100' : 'translate-x-4 opacity-0'
      }`}
    >
      {/* Header */}
      <div className="flex items-center justify-between px-4 py-3 border-b border-gray-200 bg-white/90 backdrop-blur-sm">
        <div className="flex flex-col gap-0.5">
          <div className="flex items-center gap-2">
            <div className="inline-flex h-7 w-7 items-center justify-center rounded-full bg-banana-100 text-banana-700">
              <MessageCircle size={16} />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Agent</h3>
          </div>
          <p className="text-[11px] text-gray-500 pl-9">
            用自然语言描述你想对当前PPT做的调整
          </p>
        </div>
        <button
          onClick={onClose}
          className="p-1.5 rounded-full hover:bg-gray-100 text-gray-500 hover:text-gray-800 transition-colors"
          aria-label="关闭"
        >
          <X size={16} />
        </button>
      </div>

      {/* Messages */}
      <div className="flex-1 overflow-y-auto px-4 py-3 space-y-3">
        {messages.length === 0 ? (
          <div className="mt-8 rounded-2xl bg-white/80 border border-dashed border-gray-200 px-4 py-3 text-xs text-gray-500">
            <p className="font-medium text-gray-700 mb-2">你可以这样问：</p>
            <ul className="space-y-1 list-disc list-inside">
              <li>“帮我把当前页的背景换成深色渐变”</li>
              <li>“优化这页的文案，让语气更正式一些”</li>
              <li>“根据当前大纲重写这页的要点”</li>
            </ul>
          </div>
        ) : (
          messages.map((message, index) => (
            <div
              key={index}
              className={`flex ${message.role === 'user' ? 'justify-end' : 'justify-start'}`}
            >
              <div
                className={`max-w-[80%] px-3 py-2.5 rounded-2xl text-sm leading-relaxed ${
                  message.role === 'user'
                    ? 'bg-banana-500 text-white rounded-br-sm shadow-sm'
                    : 'bg-white text-gray-800 border border-gray-200 rounded-bl-sm shadow-sm'
                }`}
              >
                <p className="whitespace-pre-wrap break-words">
                  {message.content}
                </p>
                <p
                  className={`text-xs mt-1 ${
                    message.role === 'user' ? 'text-banana-100' : 'text-gray-400'
                  }`}
                >
                  {message.timestamp.toLocaleTimeString('zh-CN', {
                    hour: '2-digit',
                    minute: '2-digit',
                  })}
                </p>
              </div>
            </div>
          ))
        )}
        {isLoading && (
          <div className="flex justify-start">
            <div className="bg-white rounded-2xl px-3 py-2 border border-gray-200 shadow-sm">
              <Loader2 size={16} className="animate-spin text-banana-500" />
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Input */}
      <div className="px-4 py-3 border-t border-gray-200 bg-white/95 backdrop-blur-sm">
        <div className="flex gap-2 items-end">
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyPress={handleKeyPress}
            placeholder="输入消息... (Enter发送, Shift+Enter换行)"
            className="flex-1 min-h-[60px] max-h-[120px] px-3 py-2 text-sm border border-gray-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-banana-500/70 focus:border-banana-400 resize-none bg-gray-50"
            disabled={isLoading}
          />
          <Button
            onClick={handleSend}
            disabled={!input.trim() || isLoading}
            variant="primary"
            className="self-end"
          >
            {isLoading ? (
              <Loader2 size={18} className="animate-spin" />
            ) : (
              <Send size={18} />
            )}
          </Button>
        </div>
      </div>
    </div>
  );
};

