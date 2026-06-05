import { useState, useRef } from "react";
import { Download, Upload, FileType, CheckCircle2, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Card, CardContent } from "@/components/ui/card";
import { toast } from "sonner";
import { useAiSuggest } from "@workspace/api-client-react";

const SELECT_CLS =
  "flex h-9 w-full items-center justify-between whitespace-nowrap rounded-md border border-input bg-transparent px-3 py-2 text-sm shadow-sm ring-offset-background focus:outline-none focus:ring-1 focus:ring-ring disabled:cursor-not-allowed disabled:opacity-50";

export function NotebookTab() {
  const [files, setFiles] = useState<File[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const triggeredHints = useRef<Set<string>>(new Set());
  const suggest = useAiSuggest();

  const [margins, setMargins] = useState("2.5");
  const [fontFamily, setFontFamily] = useState("Times New Roman");
  const [addToc, setAddToc] = useState(false);
  const [addCaptions, setAddCaptions] = useState(false);

  const fireOnce = (key: string, trigger: string, context = "") => {
    if (!triggeredHints.current.has(key)) {
      triggeredHints.current.add(key);
      suggest.mutate(
        { data: { trigger, context } },
        {
          onSuccess: (data) => {
            if (data.suggestion) {
              toast(data.suggestion, {
                action: { label: "OK", onClick: () => {} },
                cancel: { label: "Dismiss", onClick: () => {} },
                duration: 10000,
              });
            }
          },
        }
      );
    }
  };

  const handleFilesSelected = (selected: FileList | null) => {
    if (!selected) return;
    const valid = Array.from(selected).filter((f) => f.name.endsWith(".ipynb"));
    const invalid = Array.from(selected).length - valid.length;
    if (invalid > 0) toast.error(`${invalid} file(s) skipped — only .ipynb accepted`);
    if (valid.length === 0) return;

    const merged = [...files, ...valid].slice(0, 5);
    setFiles(merged);

    const totalSize = merged.reduce((s, f) => s + f.size, 0);
    if (merged.length > 1 || totalSize > 20_000) fireOnce("ipynb_large", "ipynb_large", `${merged.length} file(s), ~${Math.round(totalSize / 1024)}KB`);
    if (totalSize > 30_000) fireOnce("ipynb_images", "ipynb_images");
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    handleFilesSelected(e.dataTransfer.files);
  };

  const removeFile = (idx: number) =>
    setFiles((prev) => prev.filter((_, i) => i !== idx));

  const handleConvert = async () => {
    if (files.length === 0) {
      toast.error("Please select at least one notebook file");
      return;
    }

    setIsConverting(true);
    try {
      const formData = new FormData();
      files.forEach((f) => formData.append("files", f));
      formData.append("settings", JSON.stringify({ margins, fontFamily, addToc, addCaptions }));

      const response = await fetch("/api/convert-notebook", { method: "POST", body: formData });
      if (!response.ok) throw new Error("Conversion failed");

      const blob = await response.blob();
      const isZip = files.length > 1;
      const downloadName = isZip
        ? "notebooks.zip"
        : files[0].name.replace(".ipynb", ".docx");

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = downloadName;
      document.body.appendChild(a);
      a.click();
      URL.revokeObjectURL(url);
      document.body.removeChild(a);

      toast.success(isZip ? `${files.length} notebooks converted — downloaded as .zip` : "Notebook converted successfully");
    } catch {
      toast.error("Conversion failed — check the file is a valid .ipynb");
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <div className="flex-1 flex flex-col p-6 overflow-y-auto">
      <div className="max-w-4xl mx-auto w-full space-y-8">
        <div>
          <h2 className="text-2xl font-bold tracking-tight">Convert Jupyter Notebooks</h2>
          <p className="text-muted-foreground mt-1">
            Upload up to 5 .ipynb files. Single file downloads as .docx — multiple files download as .zip.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="md:col-span-2 space-y-4">
            {/* Drop zone */}
            <Card>
              <CardContent className="p-6">
                <div
                  className={`border-2 border-dashed rounded-lg p-10 flex flex-col items-center justify-center text-center transition-colors cursor-pointer ${isDragging ? "border-primary bg-primary/5" : "border-border hover:border-primary/50"}`}
                  onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                  onDragLeave={() => setIsDragging(false)}
                  onDrop={handleDrop}
                  onClick={() => fileInputRef.current?.click()}
                >
                  <input
                    type="file"
                    ref={fileInputRef}
                    onChange={(e) => handleFilesSelected(e.target.files)}
                    accept=".ipynb"
                    multiple
                    className="hidden"
                  />

                  <div className="w-12 h-12 bg-muted rounded-full flex items-center justify-center text-muted-foreground mb-3">
                    <Upload className="w-6 h-6" />
                  </div>
                  <p className="font-medium">Click or drag notebooks to upload</p>
                  <p className="text-sm text-muted-foreground mt-1">
                    .ipynb only — up to 5 files
                  </p>
                </div>

                {/* File list */}
                {files.length > 0 && (
                  <div className="mt-4 space-y-2">
                    {files.map((f, idx) => (
                      <div key={idx} className="flex items-center justify-between bg-muted/40 rounded-md px-3 py-2">
                        <div className="flex items-center gap-2 min-w-0">
                          <CheckCircle2 className="w-4 h-4 text-primary shrink-0" />
                          <span className="text-sm truncate">{f.name}</span>
                          <span className="text-xs text-muted-foreground shrink-0">
                            {(f.size / 1024).toFixed(1)} KB
                          </span>
                        </div>
                        <Button variant="ghost" size="icon" className="h-6 w-6 shrink-0"
                          onClick={(e) => { e.stopPropagation(); removeFile(idx); }}>
                          <X className="w-3 h-3" />
                        </Button>
                      </div>
                    ))}
                    {files.length < 5 && (
                      <p className="text-xs text-muted-foreground text-center pt-1">
                        Click the drop zone to add more (up to {5 - files.length} more)
                      </p>
                    )}
                  </div>
                )}
              </CardContent>
            </Card>

            <div className="flex justify-end">
              <Button size="lg" onClick={handleConvert} disabled={files.length === 0 || isConverting} className="gap-2">
                {isConverting ? "Converting..." : (
                  <>
                    <FileType className="w-4 h-4" />
                    Convert & Download{files.length > 1 ? " (.zip)" : ""}
                  </>
                )}
              </Button>
            </div>
          </div>

          {/* Settings panel */}
          <div className="space-y-4">
            <Card>
              <CardContent className="p-6 space-y-5">
                <h3 className="font-semibold">Settings</h3>

                <div className="space-y-1.5">
                  <Label>Font Family</Label>
                  <select className={SELECT_CLS} value={fontFamily}
                    onChange={(e) => setFontFamily(e.target.value)}>
                    <option>Times New Roman</option>
                    <option>Arial</option>
                    <option>Calibri</option>
                    <option>Georgia</option>
                    <option>Cambria</option>
                  </select>
                </div>

                <div className="space-y-1.5">
                  <Label>Margins (cm)</Label>
                  <Input type="number" step="0.1" value={margins}
                    onChange={(e) => setMargins(e.target.value)} />
                </div>

                <div className="border-t border-border pt-4 space-y-4">
                  <h4 className="text-sm font-medium">Output Options</h4>

                  <div className="space-y-1">
                    <div className="flex items-center justify-between">
                      <Label className="flex-1 cursor-pointer text-sm">Add Table of Contents</Label>
                      <Switch checked={addToc} onCheckedChange={setAddToc} />
                    </div>
                    {addToc && (
                      <p className="text-xs text-muted-foreground leading-snug">
                        Press Ctrl+A then F9 in Word to update the TOC after opening.
                      </p>
                    )}
                  </div>

                  <div className="space-y-1">
                    <div className="flex items-center justify-between">
                      <Label className="flex-1 cursor-pointer text-sm">Auto-caption images</Label>
                      <Switch checked={addCaptions} onCheckedChange={setAddCaptions} />
                    </div>
                    {addCaptions && (
                      <p className="text-xs text-muted-foreground leading-snug">
                        Adds "Figure 1", "Figure 2", etc. below each plot output.
                      </p>
                    )}
                  </div>
                </div>
              </CardContent>
            </Card>

            <Card>
              <CardContent className="p-4 space-y-2">
                <h4 className="text-sm font-medium">What gets converted</h4>
                <ul className="text-xs text-muted-foreground space-y-1">
                  <li>Markdown cells → headings and paragraphs</li>
                  <li>Code cells → shaded monospace blocks</li>
                  <li>Text outputs → light yellow blocks</li>
                  <li>Plot images → embedded pictures</li>
                  <li>DataFrames → Word tables with borders</li>
                  <li>Empty outputs → skipped silently</li>
                </ul>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </div>
  );
}
