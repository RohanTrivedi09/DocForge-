import { useState, useRef, useEffect } from "react";
import { Download, Upload, FileType, CheckCircle2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent } from "@/components/ui/card";
import { toast } from "sonner";

export function NotebookTab() {
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Settings
  const [margins, setMargins] = useState("2.5");
  const [fontFamily, setFontFamily] = useState("Times New Roman");

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile.name.endsWith(".ipynb")) {
        setFile(droppedFile);
      } else {
        toast.error("Please upload a .ipynb file");
      }
    }
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
    }
  };

  const handleConvert = async () => {
    if (!file) {
      toast.error("Please select a notebook file first");
      return;
    }

    setIsConverting(true);
    
    try {
      const formData = new FormData();
      formData.append("file", file);
      
      const settings = {
        margins,
        fontFamily
      };
      formData.append("settings", JSON.stringify(settings));

      const response = await fetch("/api/convert-notebook", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Conversion failed");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = file.name.replace(".ipynb", ".docx");
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      toast.success("Notebook converted successfully");
    } catch (error) {
      console.error(error);
      toast.error("Failed to convert notebook");
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <div className="flex-1 flex flex-col p-6 overflow-y-auto">
      <div className="max-w-4xl mx-auto w-full space-y-8">
        <div>
          <h2 className="text-2xl font-bold tracking-tight">Convert Jupyter Notebook</h2>
          <p className="text-muted-foreground mt-1">
            Transform your .ipynb files into perfectly formatted Word documents.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="md:col-span-2 space-y-6">
            <Card>
              <CardContent className="p-6">
                <div 
                  className={`border-2 border-dashed rounded-lg p-12 flex flex-col items-center justify-center text-center transition-colors ${isDragging ? 'border-primary bg-primary/5' : 'border-border'} ${file ? 'bg-muted/50' : ''}`}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  onClick={() => fileInputRef.current?.click()}
                  style={{ cursor: "pointer" }}
                >
                  <input 
                    type="file" 
                    ref={fileInputRef} 
                    onChange={handleFileSelect} 
                    accept=".ipynb" 
                    className="hidden" 
                  />
                  
                  {file ? (
                    <div className="flex flex-col items-center space-y-3">
                      <div className="w-12 h-12 bg-primary/10 rounded-full flex items-center justify-center text-primary">
                        <CheckCircle2 className="w-6 h-6" />
                      </div>
                      <div>
                        <p className="font-medium text-foreground">{file.name}</p>
                        <p className="text-sm text-muted-foreground">{(file.size / 1024).toFixed(1)} KB</p>
                      </div>
                      <Button variant="outline" size="sm" onClick={(e) => { e.stopPropagation(); setFile(null); }} className="mt-2">
                        Change File
                      </Button>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center space-y-3">
                      <div className="w-12 h-12 bg-muted rounded-full flex items-center justify-center text-muted-foreground">
                        <Upload className="w-6 h-6" />
                      </div>
                      <div>
                        <p className="font-medium text-foreground">Click or drag notebook to upload</p>
                        <p className="text-sm text-muted-foreground mt-1">Supports .ipynb files only</p>
                      </div>
                    </div>
                  )}
                </div>
              </CardContent>
            </Card>

            <div className="flex justify-end">
              <Button size="lg" onClick={handleConvert} disabled={!file || isConverting} className="gap-2">
                {isConverting ? (
                  <>Converting...</>
                ) : (
                  <>
                    <FileType className="w-4 h-4" />
                    Convert & Download
                  </>
                )}
              </Button>
            </div>
          </div>

          <div className="space-y-6">
            <Card>
              <CardContent className="p-6 space-y-4">
                <h3 className="font-semibold text-lg">Settings</h3>
                
                <div className="space-y-3">
                  <div className="space-y-1.5">
                    <Label htmlFor="fontFamily">Font Family</Label>
                    <select 
                      id="fontFamily"
                      className="flex h-9 w-full items-center justify-between whitespace-nowrap rounded-md border border-input bg-transparent px-3 py-2 text-sm shadow-sm ring-offset-background placeholder:text-muted-foreground focus:outline-none focus:ring-1 focus:ring-ring disabled:cursor-not-allowed disabled:opacity-50"
                      value={fontFamily}
                      onChange={(e) => setFontFamily(e.target.value)}
                    >
                      <option value="Times New Roman">Times New Roman</option>
                      <option value="Arial">Arial</option>
                      <option value="Calibri">Calibri</option>
                      <option value="Georgia">Georgia</option>
                      <option value="Cambria">Cambria</option>
                    </select>
                  </div>
                  
                  <div className="space-y-1.5">
                    <Label htmlFor="margins">Margins (cm)</Label>
                    <Input 
                      id="margins" 
                      type="number" 
                      step="0.1" 
                      value={margins} 
                      onChange={(e) => setMargins(e.target.value)} 
                    />
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </div>
  );
}