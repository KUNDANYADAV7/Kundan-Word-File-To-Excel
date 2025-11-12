
"use client";

import { useState } from "react";
import { Loader2, Download, CheckCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardFooter,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import FileUploader from "@/components/file-uploader";
import { parseFile, generateExcel, Question } from "@/lib/converter";
import { useToast } from "@/hooks/use-toast";
import { saveAs } from 'file-saver';

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [isDownloading, setIsDownloading] = useState(false);
  const { toast } = useToast();

  const handleFileSelect = (selectedFile: File | null) => {
    if (selectedFile) {
        const allowedTypes = [
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/pdf"
        ];
        if (!allowedTypes.includes(selectedFile.type)) {
            toast({
                variant: "destructive",
                title: "Invalid File Type",
                description: "Please upload a valid .docx or .pdf file.",
            });
            setFile(null);
            return;
        }
    }
    setFile(selectedFile);
  };

  const handleDownload = async () => {
    if (!file) {
      toast({
        variant: "destructive",
        title: "No file selected",
        description: "Please upload a file first.",
      });
      return;
    }
    
    setIsDownloading(true);
    
    try {
      const parsedQuestions = await parseFile(file);

      if (!parsedQuestions || parsedQuestions.length === 0) {
        toast({
          variant: "destructive",
          title: "Parsing Failed",
          description: "No questions could be extracted. Please check if the document is formatted correctly (e.g., questions numbered '1.', options labeled '(A)').",
        });
        setIsDownloading(false);
        return;
      }
      
      const excelBlob = await generateExcel(parsedQuestions);
      saveAs(excelBlob, `${file.name.replace(/\.(docx|pdf)$/, '') || 'download'}.xlsx`);
      
      toast({
        title: "Download Successful!",
        description: "Your Excel file has been downloaded.",
        action: <CheckCircle className="text-green-500" />,
      });

    } catch (error) {
      console.error(error);
      const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
      toast({
          variant: "destructive",
          title: "Operation Failed",
          description: errorMessage,
      });
    } finally {
        setIsDownloading(false);
    }
  };
  
  const FileUpIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="h-8 w-8 text-primary">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
      <polyline points="17 8 12 3 7 8"/>
      <line x1="12" x2="12" y1="3" y2="15"/>
    </svg>
  );

  return (
    <div className="flex min-h-screen w-full flex-col items-center bg-background px-4 py-8">
      <div className="absolute top-0 left-0 w-full h-full bg-primary/10 -z-10 [mask-image:radial-gradient(ellipse_at_center,white_20%,transparent_70%)]"></div>
      <main className="w-full max-w-4xl">
        <Card className="shadow-xl ring-1 ring-black/5">
          <CardHeader className="items-center text-center">
             <div className="flex h-16 w-16 items-center justify-center rounded-2xl bg-primary/10 mb-4">
                <FileUpIcon />
            </div>
            <CardTitle className="font-headline text-2xl">
              DocX/PDF to Excel Converter
            </CardTitle>
            <CardDescription>
              Upload your quiz and download the structured Excel sheet.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <FileUploader onFileSelect={handleFileSelect} file={file} />
          </CardContent>
          {file && (
             <CardFooter className="flex-col sm:flex-row gap-4">
                <Button
                    className="w-full font-bold text-lg py-6"
                    size="lg"
                    onClick={handleDownload}
                    disabled={isDownloading}
                >
                   {isDownloading ? (
                        <>
                            <Loader2 className="mr-2 h-5 w-5 animate-spin" />
                            Processing...
                        </>
                    ) : (
                        <>
                            <Download className="mr-2 h-5 w-5" />
                            Download Excel File
                        </>
                    )}
                </Button>
            </CardFooter>
          )}
        </Card>

        <footer className="mt-8 text-center text-sm text-muted-foreground">
            <p className="font-semibold">100% Private and Secure</p>
            <p>All processing happens locally in your browser. No data is ever uploaded to a server.</p>
        </footer>
      </main>
    </div>
  );
}
