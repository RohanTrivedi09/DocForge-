import { useState, useRef } from "react";
import { FileType, Upload, X, CheckCircle2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { toast } from "sonner";
import { useAiSuggest } from "@workspace/api-client-react";

const SELECT_CLS =
  "flex h-9 w-full items-center justify-between whitespace-nowrap rounded-md border border-input bg-transparent px-3 py-2 text-sm shadow-sm ring-offset-background focus:outline-none focus:ring-1 focus:ring-ring disabled:cursor-not-allowed disabled:opacity-50";

export function NotebookTab() {
  const [files, setFiles]         = useState<File[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const triggeredHints = useRef<Set<string>>(new Set());
  const suggest = useAiSuggest();

  // Document info
  const [courseCode, setCourseCode] = useState("");
  const [subject, setSubject]       = useState("");
  const [labNumber, setLabNumber]   = useState("Lab-1");
  const [enrollNo, setEnrollNo]     = useState("");

  // Options
  const [syntaxColor,    setSyntaxColor]    = useState(true);
  const [embedImages,    setEmbedImages]    = useState(true);
  const [showDataFrames, setShowDataFrames] = useState(true);
  const [addToc,         setAddToc]         = useState(false);
  const [showWarnings,   setShowWarnings]   = useState(false);
  const [pageNumbers,    setPageNumbers]    = useState(false);
  const [pagePos,        setPagePos]        = useState<"center" | "right">("center");

  // Output filename
  const [filename, setFilename] = useState("notebook.docx");

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
    const valid   = Array.from(selected).filter((f) => f.name.endsWith(".ipynb"));
    const skipped = Array.from(selected).length - valid.length;
    if (skipped > 0) toast.error(`${skipped} file(s) skipped — only .ipynb accepted`);
    if (!valid.length) return;

    const merged = [...files, ...valid].slice(0, 5);
    setFiles(merged);

    // Auto-fill filename from first file
    if (files.length === 0 && valid.length === 1) {
      setFilename(valid[0].name.replace(".ipynb", ".docx"));
    }

    // AI triggers
    const totalSize = merged.reduce((s, f) => s + f.size, 0);
    if (merged.length > 1 || totalSize > 20_000) {
      fireOnce("ipynb_large", "ipynb_large", `${merged.length} file(s), ~${Math.round(totalSize / 1024)} KB`);
    }
    if (totalSize > 60_000) {
      fireOnce("ipynb_images", "ipynb_images");
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    handleFilesSelected(e.dataTransfer.files);
  };

  const removeFile = (idx: number) => setFiles((p) => p.filter((_, i) => i !== idx));

  const handleConvert = async () => {
    if (!files.length) { toast.error("Upload at least one .ipynb file first"); return; }

    setIsConverting(true);
    try {
      const settings = {
        courseCode, subject, labNumber, enrollNo,
        syntaxColor, embedImages, showDataFrames, addToc,
        showWarnings, pageNumbers, pageNumberPos: pagePos,
        addCaptions: true,
        margins: 2.5,
        filename,
      };
      const formData = new FormData();
      files.forEach((f) => formData.append("files", f));
      formData.append("settings", JSON.stringify(settings));

      const response = await fetch("/api/convert-notebook", { method: "POST", body: formData });
      if (!response.ok) throw new Error("Conversion failed");

      const blob = await response.blob();
      const isZip = files.length > 1;
      const dlName = isZip ? "notebooks.zip" : (filename.endsWith(".docx") ? filename : filename + ".docx");

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = dlName;
      document.body.appendChild(a); a.click();
      URL.revokeObjectURL(url); document.body.removeChild(a);

      toast.success(isZip
        ? `${files.length} notebooks converted — downloaded as .zip`
        : "Notebook converted successfully");
    } catch {
      toast.error("Conversion failed — check the file is a valid .ipynb");
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <div className="flex-1 overflow-y-auto bg-muted/20">
      <div className="max-w-3xl mx-auto p-6 space-y-6">

        {/* ── Title ── */}
        <div>
          <h2 className="text-xl font-bold tracking-tight flex items-center gap-2">
            📓 Notebook to Word Converter
          </h2>
          <p className="text-sm text-muted-foreground mt-1">
            Upload up to 5 .ipynb files. Single file → .docx · Multiple files → .zip
          </p>
        </div>

        {/* ── Drop zone ── */}
        <div
          className={`border-2 border-dashed rounded-xl p-10 flex flex-col items-center justify-center text-center transition-colors cursor-pointer ${isDragging ? "border-primary bg-primary/5" : "border-border hover:border-primary/40 hover:bg-muted/30"}`}
          onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <input type="file" ref={fileInputRef} accept=".ipynb" multiple className="hidden"
            onChange={(e) => handleFilesSelected(e.target.files)} />
          <div className="w-12 h-12 bg-muted rounded-full flex items-center justify-center mb-3">
            <Upload className="w-6 h-6 text-muted-foreground" />
          </div>
          <p className="font-medium">Drop .ipynb here or click to upload</p>
          <p className="text-sm text-muted-foreground mt-1">.ipynb only — up to 5 files</p>
        </div>

        {/* ── File list ── */}
        {files.length > 0 && (
          <div className="space-y-2">
            {files.map((f, idx) => (
              <div key={idx} className="flex items-center gap-2 bg-muted/50 rounded-lg px-3 py-2">
                <CheckCircle2 className="w-4 h-4 text-primary shrink-0" />
                <span className="text-sm flex-1 truncate">{f.name}</span>
                <span className="text-xs text-muted-foreground shrink-0">{(f.size / 1024).toFixed(1)} KB</span>
                <Button variant="ghost" size="icon" className="h-6 w-6 shrink-0"
                  onClick={() => removeFile(idx)}>
                  <X className="w-3 h-3" />
                </Button>
              </div>
            ))}
          </div>
        )}

        {/* ── Document Info ── */}
        <div className="rounded-xl border border-border bg-card p-5 space-y-4">
          <h3 className="text-sm font-semibold text-muted-foreground uppercase tracking-wider">
            Document Info
          </h3>
          <div className="grid grid-cols-2 gap-3">
            <div className="space-y-1">
              <Label className="text-sm">Course Code</Label>
              <Input placeholder="e.g. ECSCI24305" value={courseCode}
                onChange={(e) => setCourseCode(e.target.value)} />
            </div>
            <div className="space-y-1">
              <Label className="text-sm">Subject</Label>
              <Input placeholder="e.g. Deep Learning" value={subject}
                onChange={(e) => setSubject(e.target.value)} />
            </div>
            <div className="space-y-1">
              <Label className="text-sm">Lab Number</Label>
              <Input placeholder="e.g. Lab-7" value={labNumber}
                onChange={(e) => setLabNumber(e.target.value)} />
            </div>
            <div className="space-y-1">
              <Label className="text-sm">Enrollment No.</Label>
              <Input placeholder="e.g. 1AUA23BCS159" value={enrollNo}
                onChange={(e) => setEnrollNo(e.target.value)} />
            </div>
          </div>
          <p className="text-xs text-muted-foreground">
            These appear in the header (Course Code · Subject) and footer (Lab Number · Enrollment No.) on every page.
          </p>
        </div>

        {/* ── Options ── */}
        <div className="rounded-xl border border-border bg-card p-5 space-y-4">
          <h3 className="text-sm font-semibold text-muted-foreground uppercase tracking-wider">Options</h3>

          <div className="space-y-3">
            {[
              { label: "Syntax-colored code blocks", desc: "Keywords blue, strings red, comments green", val: syntaxColor, set: setSyntaxColor },
              { label: "Embed plot images", desc: "Base64 PNG/JPEG outputs embedded at 12 cm wide", val: embedImages, set: setEmbedImages },
              { label: "Render DataFrames as tables", desc: "Detected tabular text → Word table with borders", val: showDataFrames, set: setShowDataFrames },
              { label: "Add Table of Contents", desc: "Press Ctrl+A → F9 in Word to update it", val: addToc, set: setAddToc },
              { label: "Show warnings / stderr", desc: "Stderr output shown in yellow box (hidden by default)", val: showWarnings, set: setShowWarnings },
            ].map(({ label, desc, val, set }) => (
              <div key={label} className="flex items-start justify-between gap-4">
                <div className="flex-1">
                  <p className="text-sm font-medium leading-none">{label}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">{desc}</p>
                </div>
                <Switch checked={val} onCheckedChange={set} />
              </div>
            ))}
          </div>

          <div className="border-t border-border pt-3 space-y-3">
            <div className="flex items-start justify-between gap-4">
              <div className="flex-1">
                <p className="text-sm font-medium leading-none">Page numbers</p>
                <p className="text-xs text-muted-foreground mt-0.5">Shown in footer alongside lab/enrollment info</p>
              </div>
              <Switch checked={pageNumbers} onCheckedChange={setPageNumbers} />
            </div>
            {pageNumbers && (
              <div className="flex items-center gap-3 pl-1">
                <Label className="text-sm text-muted-foreground shrink-0">Position</Label>
                <div className="flex gap-2">
                  {(["center", "right"] as const).map((pos) => (
                    <Button key={pos} size="sm" variant={pagePos === pos ? "secondary" : "outline"}
                      className="h-7 px-3 text-xs capitalize" onClick={() => setPagePos(pos)}>
                      {pos}
                    </Button>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>

        {/* ── Filename + Convert ── */}
        <div className="rounded-xl border border-border bg-card p-5 space-y-3">
          <div className="flex items-center gap-3">
            <Label className="text-sm shrink-0">Filename</Label>
            <div className="flex-1 flex items-center gap-1">
              <Input value={filename} onChange={(e) => setFilename(e.target.value)}
                placeholder="notebook.docx" className="flex-1" />
              <span className="text-sm text-muted-foreground shrink-0">
                {files.length > 1 ? "(batch → .zip)" : ".docx"}
              </span>
            </div>
          </div>

          <Button size="lg" className="w-full gap-2 font-semibold"
            onClick={handleConvert} disabled={!files.length || isConverting}>
            {isConverting ? "Converting…" : (
              <>
                <FileType className="w-4 h-4" />
                Convert & Download{files.length > 1 ? " (.zip)" : ""}
              </>
            )}
          </Button>
        </div>

        {/* ── What gets converted ── */}
        <div className="rounded-xl border border-border bg-muted/30 p-4">
          <h4 className="text-sm font-medium mb-2">What gets converted</h4>
          <ul className="text-xs text-muted-foreground space-y-1 grid grid-cols-2 gap-x-4">
            <li>Markdown → headings & paragraphs</li>
            <li>Code → syntax-coloured gray blocks</li>
            <li>Text outputs → indented Courier text</li>
            <li>Plots → embedded 12 cm images</li>
            <li>DataFrames → Word tables</li>
            <li>Stderr → yellow box (if enabled)</li>
            <li>Errors → red Courier text</li>
            <li>Empty outputs → skipped</li>
          </ul>
        </div>

      </div>
    </div>
  );
}
