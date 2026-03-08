/**
 * Inline Predictive Textarea Component
 * Shows predictive text completion inline like modern IDEs
 */

import React, { useState, useRef, useEffect } from 'react';
import { predictiveText } from '@/lib/predictiveText';

interface InlineTextareaProps {
  value: string;
  onChange: (value: string) => void;
  onBlur?: () => void;
  placeholder?: string;
  rows?: number;
  className?: string;
  context: string;
  disabled?: boolean;
}

export function InlineTextarea({
  value,
  onChange,
  onBlur,
  placeholder,
  rows = 3,
  className = '',
  context,
  disabled = false
}: InlineTextareaProps) {
  const [completion, setCompletion] = useState<string>('');
  const [showCompletion, setShowCompletion] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const ghostRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    if (value.length < 2) {
      setCompletion('');
      setShowCompletion(false);
      return;
    }

    const timer = setTimeout(() => {
      const inline = predictiveText.getInlineCompletion(value, context);
      if (inline) {
        setCompletion(inline);
        setShowCompletion(true);
      } else {
        setCompletion('');
        setShowCompletion(false);
      }
    }, 300); // Short delay for responsiveness

    return () => clearTimeout(timer);
  }, [value, context]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Tab' && showCompletion && completion) {
      e.preventDefault();
      const newValue = value + completion;
      onChange(newValue);
      setCompletion('');
      setShowCompletion(false);
    } else if (e.key === 'Escape' && showCompletion) {
      e.preventDefault();
      setCompletion('');
      setShowCompletion(false);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    onChange(e.target.value);
  };

  const handleBlur = () => {
    if (value.trim().length > 5) {
      predictiveText.learnFromInput(value, context);
    }
    setShowCompletion(false);
    if (onBlur) {
      onBlur();
    }
  };

  return (
    <div className="relative">
      {/* Ghost textarea for completion preview */}
      {showCompletion && completion && (
        <textarea
          ref={ghostRef}
          value={value + completion}
          readOnly
          tabIndex={-1}
          rows={rows}
          className={`absolute inset-0 resize-none pointer-events-none text-gray-400 bg-transparent border-transparent ${className}`}
          style={{
            color: 'rgba(156, 163, 175, 0.6)', // Tailwind gray-400 with transparency
            caretColor: 'transparent',
            overflow: 'hidden'
          }}
        />
      )}
      
      {/* Actual textarea */}
      <textarea
        ref={textareaRef}
        value={value}
        onChange={handleChange}
        onKeyDown={handleKeyDown}
        onBlur={handleBlur}
        placeholder={placeholder}
        rows={rows}
        disabled={disabled}
        className={`relative bg-transparent ${className}`}
        style={{
          resize: 'none'
        }}
      />
      
      {/* Completion hint */}
      {showCompletion && completion && (
        <div className="absolute -bottom-6 left-0 text-xs text-gray-500">
          Press Tab to complete "{completion.trim()}"
        </div>
      )}
    </div>
  );
}