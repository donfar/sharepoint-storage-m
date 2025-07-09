import { useState } from 'react';
import { Button } from "@/components/ui/button";
import { Copy, Check } from "@phosphor-icons/react";

interface CodeBlockProps {
  code: string;
  language?: string;
  showLineNumbers?: boolean;
}

export function CodeBlock({ code, language = 'powershell', showLineNumbers = false }: CodeBlockProps) {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    await navigator.clipboard.writeText(code);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="code-block relative rounded-md bg-muted">
      <Button
        variant="ghost"
        size="sm"
        onClick={handleCopy}
        className="copy-button absolute top-2 right-2 h-8 w-8 p-0"
        aria-label="Copy code"
      >
        {copied ? (
          <Check className="h-4 w-4 text-green-500" weight="bold" />
        ) : (
          <Copy className="h-4 w-4" />
        )}
      </Button>
      <pre className={`p-4 overflow-x-auto ${showLineNumbers ? 'line-numbers' : ''}`}>
        <code>{code}</code>
      </pre>
    </div>
  );
}