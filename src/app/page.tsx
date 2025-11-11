"use client";

import { useState } from "react";
import { Loader2, FileDown, CheckCircle } from "lucide-react";
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
import { convertDocxToExcel } from "@/lib/converter";
import { useToast } from "@/hooks/use-toast";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [isConverting, setIsConverting] = useState(false);
  const { toast } = useToast();

  const handleFileSelect = (selectedFile: File | null) => {
    if (
      selectedFile &&
      selectedFile.type !==
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ) {
      toast({
        variant: "destructive",
        title: "Invalid File Type",
        description: "Please upload a valid .docx file.",
      });
      setFile(null);
      return;
    }
    setFile(selectedFile);
  };

  const handleConvert = async () => {
    if (!file) return;

    setIsConverting(true);
    try {
      await convertDocxToExcel(file);
      toast({
        title: "Success!",
        description: "Your Excel file has been generated and downloaded.",
        action: <CheckCircle className="text-green-500" />,
      });
    } catch (error) {
      console.error(error);
      const errorMessage = error instanceof Error ? error.message : "An unknown error occurred.";
      toast({
        variant: "destructive",
        title: "Conversion Failed",
        description: errorMessage,
      });
    } finally {
      setIsConverting(false);
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
    <div className="flex min-h-screen w-full items-center justify-center bg-background px-4">
      <div className="absolute top-0 left-0 w-full h-full bg-primary/10 -z-10 [mask-image:radial-gradient(ellipse_at_center,white_20%,transparent_70%)]"></div>
      <main className="w-full max-w-lg">
        <Card className="shadow-xl ring-1 ring-black/5">
          <CardHeader className="items-center text-center">
             <div className="flex h-16 w-16 items-center justify-center rounded-2xl bg-primary/10 mb-4">
                <FileUpIcon />
            </div>
            <CardTitle className="font-headline text-2xl">
              DocX to Excel Converter
            </CardTitle>
            <CardDescription>
              Upload your quiz in .docx format to convert it into a structured Excel sheet, complete with images.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <FileUploader onFileSelect={handleFileSelect} file={file} />
          </CardContent>
          <CardFooter>
            <Button
              className="w-full font-bold text-lg py-6 bg-accent text-accent-foreground hover:bg-accent/90 focus-visible:ring-accent-foreground/50"
              size="lg"
              onClick={handleConvert}
              disabled={!file || isConverting}
            >
              {isConverting ? (
                <>
                  <Loader2 className="mr-2 h-5 w-5 animate-spin" />
                  Converting...
                </>
              ) : (
                <>
                  <FileDown className="mr-2 h-5 w-5" />
                  Convert & Download
                </>
              )}
            </Button>
          </CardFooter>
        </Card>
        <footer className="mt-8 text-center text-sm text-muted-foreground">
            <p className="font-semibold">100% Private and Secure</p>
            <p>All processing happens locally in your browser. No data is ever uploaded to a server.</p>
        </footer>
      </main>
    </div>
  );
}
