"use client";

import { useRef, useState, type DragEvent, type ChangeEvent } from "react";
import { UploadCloud, FileText, X } from "lucide-react";
import { cn } from "@/lib/utils";

interface FileUploaderProps {
  onFileSelect: (file: File | null) => void;
  file: File | null;
}

export default function FileUploader({ onFileSelect, file }: FileUploaderProps) {
  const [isDragging, setIsDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setIsDragging(true);
    } else if (e.type === "dragleave") {
      setIsDragging(false);
    }
  };

  const handleDrop = (e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      onFileSelect(e.dataTransfer.files[0]);
    }
  };
  
  const handleChange = (e: ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    if (e.target.files && e.target.files[0]) {
      onFileSelect(e.target.files[0]);
    }
  };

  const onButtonClick = () => {
    inputRef.current?.click();
  };
  
  const onRemoveFile = () => {
      onFileSelect(null);
      if(inputRef.current) {
        inputRef.current.value = "";
      }
  }

  if (file) {
    return (
      <div className="flex items-center justify-between rounded-lg border bg-secondary/50 p-4 transition-all">
        <div className="flex items-center gap-3 overflow-hidden">
          <FileText className="h-6 w-6 shrink-0 text-primary" />
          <span className="truncate text-sm font-medium text-foreground">{file.name}</span>
        </div>
        <button 
            onClick={onRemoveFile} 
            className="group rounded-full p-1 text-muted-foreground transition-colors hover:bg-destructive/10 hover:text-destructive"
            aria-label="Remove file"
        >
          <X className="h-5 w-5" />
        </button>
      </div>
    );
  }

  return (
    <div
      className={cn(
        "flex w-full cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed border-border bg-background p-10 text-center transition-colors duration-300",
        isDragging ? "border-primary bg-primary/10" : "hover:border-border/80 hover:bg-muted/50"
      )}
      onDragEnter={handleDrag}
      onDragLeave={handleDrag}
      onDragOver={handleDrag}
      onDrop={handleDrop}
      onClick={onButtonClick}
      role="button"
      aria-label="File upload zone"
    >
      <input
        ref={inputRef}
        type="file"
        id="file-upload"
        className="hidden"
        onChange={handleChange}
        accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      />
      <UploadCloud className="mb-4 h-12 w-12 text-muted-foreground transition-transform group-hover:scale-110" />
      <p className="font-semibold text-foreground">Drag & drop your .docx file</p>
      <p className="text-sm text-muted-foreground">or click to browse</p>
    </div>
  );
}
