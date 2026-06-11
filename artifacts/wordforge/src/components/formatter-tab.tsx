import { useState, useRef, useEffect } from "react";
import { Download, Upload, Eye, Settings, AlignLeft, AlignCenter, AlignRight, FileDown, FileUp } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { ResizablePanelGroup, ResizablePanel, ResizableHandle } from "@/components/ui/resizable";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { toast } from "sonner";
import { useAiSuggest } from "@workspace/api-client-react";
import { ProfileSection, loadProfile } from "@/components/profile-section";

const SELECT_CLS =
  "flex h-9 w-full items-center justify-between whitespace-nowrap rounded-md border border-input bg-transparent px-3 py-2 text-sm shadow-sm ring-offset-background focus:outline-none focus:ring-1 focus:ring-ring disabled:cursor-not-allowed disabled:opacity-50";

const PRESETS = {
  "GTU Standard": {
    fontFamily: "Times New Roman", bodySize: 12, bodyColor: "#000000",
    h1Size: 16, h1Color: "#000000", h2Size: 14, h2Color: "#000000",
    h3Size: 12, h3Color: "#000000", h4Size: 11, h4Color: "#000000",
    lineSpacing: "1.5", spacingBefore: 0, spacingAfter: 6,
    marginTop: 2.5, marginBottom: 2.5, marginLeft: 2.5, marginRight: 2.5,
  },
  "Adani University": {
    fontFamily: "Times New Roman", bodySize: 12, bodyColor: "#000000",
    h1Size: 16, h1Color: "#1F3864", h2Size: 14, h2Color: "#2E74B5",
    h3Size: 12, h3Color: "#000000", h4Size: 11, h4Color: "#000000",
    lineSpacing: "1.5", spacingBefore: 0, spacingAfter: 6,
    marginTop: 2.5, marginBottom: 2.5, marginLeft: 2.5, marginRight: 2.5,
  },
  "IEEE": {
    fontFamily: "Times New Roman", bodySize: 10, bodyColor: "#000000",
    h1Size: 12, h1Color: "#000000", h2Size: 11, h2Color: "#000000",
    h3Size: 10, h3Color: "#000000", h4Size: 10, h4Color: "#000000",
    lineSpacing: "1.0", spacingBefore: 0, spacingAfter: 4,
    marginTop: 2.0, marginBottom: 2.0, marginLeft: 2.0, marginRight: 2.0,
  },
  "APA 7th": {
    fontFamily: "Times New Roman", bodySize: 12, bodyColor: "#000000",
    h1Size: 14, h1Color: "#000000", h2Size: 13, h2Color: "#000000",
    h3Size: 12, h3Color: "#000000", h4Size: 12, h4Color: "#000000",
    lineSpacing: "2.0", spacingBefore: 0, spacingAfter: 0,
    marginTop: 2.54, marginBottom: 2.54, marginLeft: 2.54, marginRight: 2.54,
  },
  "Gujarat Govt. SOP": {
    fontFamily: "Times New Roman", bodySize: 12, bodyColor: "#000000",
    h1Size: 14, h1Color: "#000080", h2Size: 14, h2Color: "#C00000",
    h3Size: 12, h3Color: "#1F4D78", h4Size: 12, h4Color: "#1F4D78",
    lineSpacing: "1.5", spacingBefore: 0, spacingAfter: 6,
    marginTop: 2.54, marginBottom: 2.54, marginLeft: 2.54, marginRight: 2.54,
  },
} as const;

type SettingsKey =
  | "fontFamily" | "bodySize" | "bodyColor"
  | "h1Size" | "h1Color" | "h2Size" | "h2Color"
  | "h3Size" | "h3Color" | "h4Size" | "h4Color"
  | "lineSpacing" | "spacingBefore" | "spacingAfter"
  | "marginTop" | "marginBottom" | "marginLeft" | "marginRight"
  | "headerText" | "footerText" | "pageNumbers" | "pageNumberPos"
  | "pageXofY" | "diffFirstPage" | "diffOddEven";

const SETTINGS_STORAGE_KEY = "docforge_last_settings";

const DEFAULT_SETTINGS = {
  fontFamily: "Times New Roman",
  bodySize: 12,
  bodyColor: "#000000",
  h1Size: 16,
  h1Color: "#1F3864",
  h2Size: 14,
  h2Color: "#2E74B5",
  h3Size: 12,
  h3Color: "#000000",
  h4Size: 11,
  h4Color: "#000000",
  lineSpacing: "1.5",
  spacingBefore: 0,
  spacingAfter: 6,
  marginTop: 2.5,
  marginBottom: 2.5,
  marginLeft: 2.5,
  marginRight: 2.5,
  headerText: "",
  footerText: "",
  pageNumbers: true,
  pageNumberPos: "center",
  pageXofY: false,
  diffFirstPage: false,
  diffOddEven: false,
};

function loadLastSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (raw) return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch {}
  return DEFAULT_SETTINGS;
}

export function FormatterTab() {
  const [file, setFile] = useState<File | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const importRef = useRef<HTMLInputElement>(null);
  const suggest = useAiSuggest();
  const triggeredHints = useRef<Set<string>>(new Set());

  const [settings, setSettings] = useState(loadLastSettings);

  const [filename, setFilename] = useState("submission.docx");

  const [coverPage, setCoverPage] = useState({
    enabled: false,
    title: "",
    studentName: "",
    rollNumber: "",
    subject: "",
    department: "",
    universityName: "",
    date: "",
  });

  const [isFormatting, setIsFormatting] = useState(false);

  // Gujarat Govt. SOP specific state
  const [activePreset, setActivePreset] = useState<string | null>(null);
  const [sopChapterNumber, setSopChapterNumber] = useState("1");
  const [sopChapterTitle, setSopChapterTitle] = useState("");
  const [sopContent, setSopContent] = useState("");
  const [sopRefFile, setSopRefFile] = useState<File | null>(null);
  const sopRefInputRef = useRef<HTMLInputElement>(null);
  const isSOP = activePreset === "Gujarat Govt. SOP";

  // Auto-fill cover page from profile on mount
  useEffect(() => {
    const profile = loadProfile();
    if (profile) {
      setCoverPage((p) => ({
        ...p,
        studentName: p.studentName || profile.name,
        universityName: p.universityName || profile.university,
        department: p.department || profile.department,
      }));
    }
  }, []);

  const triggerHint = (trigger: string, context = "") => {
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
  };

  const fireOnce = (key: string, trigger: string, context = "") => {
    if (!triggeredHints.current.has(key)) {
      triggeredHints.current.add(key);
      triggerHint(trigger, context);
    }
  };

  const updateSetting = (key: SettingsKey, value: any) => {
    setSettings((prev: typeof DEFAULT_SETTINGS) => {
      const next = { ...prev, [key]: value };

      if (key === "h1Size" && value > 20) fireOnce("h1_large", "h1_large", `H1 size: ${value}pt`);
      if (key === "lineSpacing" && value === "2.0") fireOnce("double_spacing", "double_spacing");

      if ((key === "h1Size" || key === "h2Size")) {
        const h1 = key === "h1Size" ? value : prev.h1Size;
        const h2 = key === "h2Size" ? value : prev.h2Size;
        if (h1 === h2) fireOnce("h1_h2_same", "h1_h2_same", `H1=${h1}pt, H2=${h2}pt`);
      }

      if (key.endsWith("Color") && typeof value === "string") {
        const h = value.replace("#", "");
        if (h.length === 6) {
          const r = parseInt(h.slice(0, 2), 16);
          const g = parseInt(h.slice(2, 4), 16);
          const b = parseInt(h.slice(4, 6), 16);
          const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
          if (lum < 0.3 && lum > 0.01) fireOnce("dark_color", "dark_color", `Color: ${value}`);
        }
      }

      if (key === "pageNumberPos" && value === "left" && next.pageNumbers) {
        fireOnce("footer_left_pagenum", "footer_left_pagenum");
      }

      return next;
    });
  };

  const handleHeaderFocus = () => fireOnce("header_focus", "header_focus");

  const handleFormat = async () => {
    // ── Gujarat Govt. SOP path ────────────────────────────────────────────────
    if (isSOP) {
      if (!sopChapterTitle.trim()) {
        toast.error("Enter a Chapter Title before generating the SOP");
        return;
      }
      setIsFormatting(true);
      try {
        const formData = new FormData();
        formData.append("chapter_number", sopChapterNumber);
        formData.append("chapter_title", sopChapterTitle);
        formData.append("content", sopContent);
        if (sopRefFile) formData.append("reference_docx", sopRefFile);

        const response = await fetch("/api/format-sop", { method: "POST", body: formData });
        if (!response.ok) throw new Error("SOP generation failed");

        // Show validation report from response header
        const validationRaw = response.headers.get("X-SOP-Validation");
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `SOP_Chapter_${sopChapterNumber}.docx`;
        document.body.appendChild(a); a.click();
        URL.revokeObjectURL(url); document.body.removeChild(a);

        if (validationRaw) {
          try {
            const v = JSON.parse(validationRaw);
            const issues = (v.issues || []).length;
            const msg = `✅ Generated: ${v.paragraph_count} paragraphs · ${v.h1_count} H1 · ${v.h2_count} H2 · ${issues === 0 ? "0 color mismatches" : `⚠ ${issues} color issue(s)`}`;
            toast.success(msg, { duration: 8000 });
          } catch {}
        } else {
          toast.success("SOP document generated and downloaded");
        }
      } catch {
        toast.error("Failed to generate SOP document");
      } finally {
        setIsFormatting(false);
      }
      return;
    }

    // ── Standard formatter path ───────────────────────────────────────────────
    if (settings.headerText === "") fireOnce("no_header_export", "no_header_export");

    setIsFormatting(true);
    try {
      const payload = { ...settings, filename, coverPage };
      const formData = new FormData();
      if (file) formData.append("file", file);
      formData.append("settings", JSON.stringify(payload));

      const response = await fetch("/api/format-doc", { method: "POST", body: formData });
      if (!response.ok) throw new Error("Formatting failed");

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename.endsWith(".docx") ? filename : filename + ".docx";
      document.body.appendChild(a);
      a.click();
      URL.revokeObjectURL(url);
      document.body.removeChild(a);

      // Auto-save settings to localStorage on every download
      localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
      toast.success("Document formatted and downloaded");
    } catch {
      toast.error("Failed to format document");
    } finally {
      setIsFormatting(false);
    }
  };

  const applyPreset = (name: keyof typeof PRESETS) => {
    const p = PRESETS[name];
    setSettings((prev: typeof DEFAULT_SETTINGS) => ({ ...prev, ...p }));
    setActivePreset(name);
    toast.success(`Preset applied: ${name}`);
  };

  const handleExportSettings = () => {
    const json = JSON.stringify({
      fontFamily: settings.fontFamily,
      bodySize: settings.bodySize,
      bodyColor: settings.bodyColor,
      headings: {
        h1: { size: settings.h1Size, color: settings.h1Color },
        h2: { size: settings.h2Size, color: settings.h2Color },
        h3: { size: settings.h3Size, color: settings.h3Color },
        h4: { size: settings.h4Size, color: settings.h4Color },
      },
      lineSpacing: settings.lineSpacing,
      spacingBefore: settings.spacingBefore,
      spacingAfter: settings.spacingAfter,
      margins: {
        top: settings.marginTop,
        bottom: settings.marginBottom,
        left: settings.marginLeft,
        right: settings.marginRight,
      },
      headerText: settings.headerText,
      footerText: settings.footerText,
      pageNumbers: settings.pageNumbers,
      pageNumberPosition: settings.pageNumberPos,
      pageXofY: settings.pageXofY,
      differentFirstPage: settings.diffFirstPage,
      differentOddEven: settings.diffOddEven,
    }, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "docforge_settings.json";
    document.body.appendChild(a); a.click();
    URL.revokeObjectURL(url); document.body.removeChild(a);
    toast.success("Settings exported");
  };

  const handleImportSettings = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target?.result as string);
        const h = parsed.headings || {};
        const m = parsed.margins || {};
        setSettings((prev: typeof DEFAULT_SETTINGS) => ({
          ...prev,
          fontFamily: parsed.fontFamily ?? prev.fontFamily,
          bodySize: parsed.bodySize ?? prev.bodySize,
          bodyColor: parsed.bodyColor ?? prev.bodyColor,
          h1Size: h.h1?.size ?? prev.h1Size,
          h1Color: h.h1?.color ?? prev.h1Color,
          h2Size: h.h2?.size ?? prev.h2Size,
          h2Color: h.h2?.color ?? prev.h2Color,
          h3Size: h.h3?.size ?? prev.h3Size,
          h3Color: h.h3?.color ?? prev.h3Color,
          h4Size: h.h4?.size ?? prev.h4Size,
          h4Color: h.h4?.color ?? prev.h4Color,
          lineSpacing: parsed.lineSpacing ?? prev.lineSpacing,
          spacingBefore: parsed.spacingBefore ?? prev.spacingBefore,
          spacingAfter: parsed.spacingAfter ?? prev.spacingAfter,
          marginTop: m.top ?? prev.marginTop,
          marginBottom: m.bottom ?? prev.marginBottom,
          marginLeft: m.left ?? prev.marginLeft,
          marginRight: m.right ?? prev.marginRight,
          headerText: parsed.headerText ?? prev.headerText,
          footerText: parsed.footerText ?? prev.footerText,
          pageNumbers: parsed.pageNumbers ?? prev.pageNumbers,
          pageNumberPos: parsed.pageNumberPosition ?? prev.pageNumberPos,
          pageXofY: parsed.pageXofY ?? prev.pageXofY,
          diffFirstPage: parsed.differentFirstPage ?? prev.diffFirstPage,
          diffOddEven: parsed.differentOddEven ?? prev.diffOddEven,
        }));
        toast.success("Settings imported");
      } catch {
        toast.error("Invalid settings file");
      }
    };
    reader.readAsText(f);
    e.target.value = "";
  };

  return (
    <ResizablePanelGroup direction="horizontal" className="h-full flex-1 rounded-none border-none">
      {/* ── Controls Panel ── */}
      <ResizablePanel defaultSize={35} minSize={25} maxSize={52} className="bg-card flex flex-col">
        <div className="p-4 border-b border-border flex items-center gap-2 bg-muted/30">
          <Settings className="w-4 h-4 text-muted-foreground" />
          <h2 className="font-semibold text-sm">Formatting Options</h2>
        </div>

        <div className="flex-1 overflow-y-auto p-4 custom-scrollbar">

          {/* Profile */}
          <ProfileSection
            onProfileLoad={(profile) => {
              setCoverPage((p) => ({
                ...p,
                studentName: p.studentName || profile.name,
                universityName: p.universityName || profile.university,
                department: p.department || profile.department,
              }));
            }}
          />

          <Accordion type="multiple" defaultValue={["typography", "layout", "header-footer"]} className="w-full">

            {/* Source file */}
            <AccordionItem value="upload" className="border-b-0 mb-4 bg-muted/20 p-4 rounded-lg border border-border">
              <div className="space-y-3">
                <Label className="text-sm font-semibold">Source Document (Optional)</Label>
                <div className="flex gap-2">
                  <input type="file" accept=".docx" className="hidden" ref={fileInputRef}
                    onChange={(e) => e.target.files?.[0] && setFile(e.target.files[0])} />
                  <Button variant="outline" className="w-full justify-start text-muted-foreground font-normal"
                    onClick={() => fileInputRef.current?.click()}>
                    <Upload className="w-4 h-4 mr-2" />
                    {file ? file.name : "Select .docx file..."}
                  </Button>
                  {file && <Button variant="ghost" size="icon" onClick={() => setFile(null)}>&times;</Button>}
                </div>
              </div>
            </AccordionItem>

            {/* Preset templates */}
            <AccordionItem value="presets">
              <AccordionTrigger className="hover:no-underline font-medium text-sm">Preset Templates</AccordionTrigger>
              <AccordionContent className="pt-2 space-y-3">
                <div className="grid grid-cols-2 gap-2">
                  {(Object.keys(PRESETS) as (keyof typeof PRESETS)[]).map((name) => (
                    <Button key={name} variant={activePreset === name ? "secondary" : "outline"}
                      size="sm" className="h-8 text-xs"
                      onClick={() => applyPreset(name)}>
                      {name === "Gujarat Govt. SOP" ? "🏛️ " : ""}{name}
                    </Button>
                  ))}
                </div>
                <p className="text-xs text-muted-foreground">
                  One click loads all settings for that standard.
                </p>

                {/* ── SOP-specific fields ── */}
                {isSOP && (
                  <div className="mt-3 rounded-lg border border-amber-200 bg-amber-50 p-3 space-y-3">
                    <p className="text-xs font-semibold text-amber-800 uppercase tracking-wider">
                      🏛️ SOP Document Info
                    </p>
                    <div className="grid grid-cols-[80px_1fr] gap-2 items-center">
                      <Label className="text-xs">Chapter No.</Label>
                      <Input className="h-8 text-sm" value={sopChapterNumber}
                        onChange={(e) => setSopChapterNumber(e.target.value)}
                        placeholder="e.g. 6" />
                    </div>
                    <div className="grid grid-cols-[80px_1fr] gap-2 items-center">
                      <Label className="text-xs">Chapter Title</Label>
                      <Input className="h-8 text-sm" value={sopChapterTitle}
                        onChange={(e) => setSopChapterTitle(e.target.value)}
                        placeholder="e.g. Wireless Grid Operations" />
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Chapter Content</Label>
                      <textarea
                        className="w-full h-28 rounded-md border border-input bg-background px-3 py-2 text-xs font-mono resize-y focus:outline-none focus:ring-1 focus:ring-ring"
                        placeholder={"Paste your chapter text here.\n\n# H1 Heading\n## H2 Heading\n### H3 Heading\n-> Bullet item\nObjective: ...\nReference: ...\n\nRegular body text..."}
                        value={sopContent}
                        onChange={(e) => setSopContent(e.target.value)}
                      />
                      <p className="text-xs text-muted-foreground">
                        Supports: # H1/H2/H3, -&gt; bullets, Objective:, Reference:, pipe tables
                      </p>
                    </div>
                    <div className="space-y-1">
                      <Label className="text-xs">Reference .docx (optional)</Label>
                      <input type="file" accept=".docx" className="hidden" ref={sopRefInputRef}
                        onChange={(e) => setSopRefFile(e.target.files?.[0] || null)} />
                      <Button variant="outline" size="sm" className="w-full h-8 text-xs justify-start gap-2"
                        onClick={() => sopRefInputRef.current?.click()}>
                        <Upload className="w-3.5 h-3.5" />
                        {sopRefFile ? sopRefFile.name : "Upload to verify colors..."}
                      </Button>
                      {sopRefFile && (
                        <p className="text-xs text-green-700">✓ Colors will be verified from reference</p>
                      )}
                    </div>
                  </div>
                )}
              </AccordionContent>
            </AccordionItem>

            {/* Typography */}
            <AccordionItem value="typography">
              <AccordionTrigger className="hover:no-underline font-medium text-sm">Typography</AccordionTrigger>
              <AccordionContent className="pt-2 space-y-4">
                <div className="space-y-2">
                  <Label>Font Family</Label>
                  <select className={SELECT_CLS} value={settings.fontFamily}
                    onChange={(e) => updateSetting("fontFamily", e.target.value)}>
                    <option>Times New Roman</option>
                    <option>Arial</option>
                    <option>Calibri</option>
                    <option>Georgia</option>
                    <option>Cambria</option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label>Body Size (pt)</Label>
                    <Input type="number" value={settings.bodySize}
                      onChange={(e) => updateSetting("bodySize", parseInt(e.target.value))} />
                  </div>
                  <div className="space-y-2">
                    <Label>Body Color</Label>
                    <div className="flex gap-2">
                      <Input type="color" className="w-10 p-1 h-9" value={settings.bodyColor}
                        onChange={(e) => updateSetting("bodyColor", e.target.value)} />
                      <Input type="text" className="flex-1" value={settings.bodyColor}
                        onChange={(e) => updateSetting("bodyColor", e.target.value)} />
                    </div>
                  </div>
                </div>

                <div className="space-y-3 pt-1">
                  <Label className="text-xs text-muted-foreground font-semibold uppercase tracking-wider">Headings</Label>
                  {[1, 2, 3, 4].map((level) => {
                    const sizeKey = `h${level}Size` as SettingsKey;
                    const colorKey = `h${level}Color` as SettingsKey;
                    return (
                      <div key={level} className="grid grid-cols-[28px_1fr_auto] gap-2 items-center">
                        <span className="text-sm font-medium text-muted-foreground">H{level}</span>
                        <Input type="number" className="h-8" value={settings[sizeKey] as number}
                          onChange={(e) => updateSetting(sizeKey, parseInt(e.target.value))} />
                        <div className="flex items-center gap-1">
                          <Input type="color" className="w-8 h-8 p-0.5 rounded cursor-pointer"
                            value={settings[colorKey] as string}
                            onChange={(e) => updateSetting(colorKey, e.target.value)} />
                          <span className="text-xs text-muted-foreground w-16 hidden xl:inline">
                            {settings[colorKey] as string}
                          </span>
                        </div>
                      </div>
                    );
                  })}
                </div>

                <div className="space-y-2 pt-1">
                  <Label>Line Spacing</Label>
                  <select className={SELECT_CLS} value={settings.lineSpacing}
                    onChange={(e) => updateSetting("lineSpacing", e.target.value)}>
                    <option value="1.0">Single (1.0)</option>
                    <option value="1.15">1.15</option>
                    <option value="1.5">1.5</option>
                    <option value="2.0">Double (2.0)</option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label>Spacing Before (pt)</Label>
                    <Input type="number" value={settings.spacingBefore}
                      onChange={(e) => updateSetting("spacingBefore", parseInt(e.target.value))} />
                  </div>
                  <div className="space-y-2">
                    <Label>Spacing After (pt)</Label>
                    <Input type="number" value={settings.spacingAfter}
                      onChange={(e) => updateSetting("spacingAfter", parseInt(e.target.value))} />
                  </div>
                </div>
              </AccordionContent>
            </AccordionItem>

            {/* Page Layout */}
            <AccordionItem value="layout">
              <AccordionTrigger className="hover:no-underline font-medium text-sm">Page Layout</AccordionTrigger>
              <AccordionContent className="pt-2 space-y-3">
                <Label className="text-xs text-muted-foreground font-semibold uppercase tracking-wider">Margins (cm)</Label>
                <div className="grid grid-cols-2 gap-3">
                  {(["Top", "Bottom", "Left", "Right"] as const).map((side) => {
                    const key = `margin${side}` as SettingsKey;
                    return (
                      <div key={side} className="space-y-1">
                        <Label className="text-xs">{side}</Label>
                        <Input type="number" step="0.1" value={settings[key] as number}
                          onChange={(e) => updateSetting(key, parseFloat(e.target.value))} />
                      </div>
                    );
                  })}
                </div>
              </AccordionContent>
            </AccordionItem>

            {/* Headers & Footers */}
            <AccordionItem value="header-footer">
              <AccordionTrigger className="hover:no-underline font-medium text-sm">Headers & Footers</AccordionTrigger>
              <AccordionContent className="pt-2 space-y-4">
                <div className="space-y-2">
                  <Label>Header Text</Label>
                  <Input placeholder="E.g., Department of Computer Science"
                    value={settings.headerText}
                    onChange={(e) => updateSetting("headerText", e.target.value)}
                    onFocus={handleHeaderFocus} />
                </div>
                <div className="space-y-2">
                  <Label>Footer Text</Label>
                  <Input placeholder="E.g., Confidential"
                    value={settings.footerText}
                    onChange={(e) => updateSetting("footerText", e.target.value)} />
                </div>

                <div className="space-y-3 pt-1">
                  <div className="flex items-center justify-between">
                    <Label className="flex-1 cursor-pointer">Page Numbers</Label>
                    <Switch checked={settings.pageNumbers}
                      onCheckedChange={(c) => updateSetting("pageNumbers", c)} />
                  </div>

                  {settings.pageNumbers && (
                    <>
                      <div className="flex gap-1 p-1 bg-muted rounded-md w-max">
                        {(["left", "center", "right"] as const).map((pos) => (
                          <Button key={pos}
                            variant={settings.pageNumberPos === pos ? "secondary" : "ghost"}
                            size="sm" className="h-7 px-2"
                            onClick={() => updateSetting("pageNumberPos", pos)}>
                            {pos === "left" ? <AlignLeft className="w-4 h-4" /> :
                              pos === "center" ? <AlignCenter className="w-4 h-4" /> :
                                <AlignRight className="w-4 h-4" />}
                          </Button>
                        ))}
                      </div>
                      <div className="flex items-center justify-between">
                        <Label className="flex-1 cursor-pointer text-sm">
                          "Page X of Y" format
                          <span className="block text-xs text-muted-foreground font-normal">
                            Instead of just the page number
                          </span>
                        </Label>
                        <Switch checked={settings.pageXofY}
                          onCheckedChange={(c) => updateSetting("pageXofY", c)} />
                      </div>
                    </>
                  )}

                  <div className="flex items-center justify-between">
                    <Label className="flex-1 cursor-pointer text-sm">Different First Page</Label>
                    <Switch checked={settings.diffFirstPage}
                      onCheckedChange={(c) => updateSetting("diffFirstPage", c)} />
                  </div>
                  <div className="flex items-center justify-between">
                    <Label className="flex-1 cursor-pointer text-sm">Different Odd & Even Pages</Label>
                    <Switch checked={settings.diffOddEven}
                      onCheckedChange={(c) => updateSetting("diffOddEven", c)} />
                  </div>
                </div>
              </AccordionContent>
            </AccordionItem>

            {/* Cover Page */}
            <AccordionItem value="cover-page">
              <AccordionTrigger className="hover:no-underline font-medium text-sm">Cover Page</AccordionTrigger>
              <AccordionContent className="pt-2 space-y-4">
                <div className="flex items-center justify-between">
                  <Label className="flex-1 cursor-pointer">Generate cover page</Label>
                  <Switch checked={coverPage.enabled}
                    onCheckedChange={(c) => setCoverPage((p) => ({ ...p, enabled: c }))} />
                </div>

                {coverPage.enabled && (
                  <div className="space-y-3 pt-1">
                    {([
                      ["title", "Title"],
                      ["studentName", "Student Name"],
                      ["rollNumber", "Roll Number"],
                      ["subject", "Subject"],
                      ["department", "Department"],
                      ["universityName", "University Name"],
                      ["date", "Date"],
                    ] as [keyof typeof coverPage, string][]).map(([key, label]) => (
                      <div key={key} className="space-y-1">
                        <Label className="text-sm">{label}</Label>
                        <Input value={coverPage[key] as string}
                          onChange={(e) => setCoverPage((p) => ({ ...p, [key]: e.target.value }))}
                          placeholder={label} />
                      </div>
                    ))}
                  </div>
                )}
              </AccordionContent>
            </AccordionItem>

          </Accordion>
        </div>

        {/* Bottom bar: filename + download */}
        <div className="p-4 border-t border-border bg-card space-y-3">
          <div className="flex items-center gap-2">
            <Label className="text-xs text-muted-foreground shrink-0">Filename</Label>
            <Input
              value={filename}
              onChange={(e) => setFilename(e.target.value)}
              className="h-8 text-sm"
              placeholder="submission.docx"
            />
          </div>
          <Button className="w-full font-semibold gap-2" size="lg" onClick={handleFormat} disabled={isFormatting}>
            {isFormatting ? "Formatting..." : <><Download className="w-4 h-4" />Format & Download .docx</>}
          </Button>
          <div className="flex gap-2">
            <input type="file" accept=".json" className="hidden" ref={importRef} onChange={handleImportSettings} />
            <Button variant="outline" size="sm" className="flex-1 gap-1.5 text-xs h-8"
              onClick={() => importRef.current?.click()}>
              <FileUp className="w-3.5 h-3.5" /> Import Settings
            </Button>
            <Button variant="outline" size="sm" className="flex-1 gap-1.5 text-xs h-8"
              onClick={handleExportSettings}>
              <FileDown className="w-3.5 h-3.5" /> Export Settings
            </Button>
          </div>
        </div>
      </ResizablePanel>

      <ResizableHandle withHandle />

      {/* ── Preview Panel ── */}
      <ResizablePanel defaultSize={65} className="bg-muted/30 flex flex-col relative overflow-hidden">
        <div className="absolute top-4 left-4 z-10">
          <div className="bg-background/80 backdrop-blur-sm px-3 py-1.5 rounded-full text-xs font-medium border border-border shadow-sm flex items-center gap-2">
            <Eye className="w-3 h-3 text-muted-foreground" />
            Live Preview
          </div>
        </div>

        <div className="flex-1 overflow-auto p-8 flex flex-col items-center gap-3 custom-scrollbar">
          <div
            className="bg-white shadow-xl transition-all duration-200 relative border border-gray-200"
            style={{
              width: "21cm",
              minHeight: "29.7cm",
              paddingTop: `${settings.marginTop}cm`,
              paddingBottom: `${settings.marginBottom}cm`,
              paddingLeft: `${settings.marginLeft}cm`,
              paddingRight: `${settings.marginRight}cm`,
              fontFamily: settings.fontFamily,
              color: settings.bodyColor,
              lineHeight: settings.lineSpacing,
            }}
          >
            {settings.headerText && (
              <div className="absolute top-0 left-0 right-0 border-b border-gray-100 flex items-end px-10 pb-1.5 text-gray-400 text-xs"
                style={{ height: `${settings.marginTop}cm` }}>
                {settings.headerText}
              </div>
            )}

            {settings.pageNumbers && (
              <div className={`absolute bottom-4 left-0 right-0 px-10 text-gray-400 text-xs flex ${settings.pageNumberPos === "left" ? "justify-start" : settings.pageNumberPos === "right" ? "justify-end" : "justify-center"}`}>
                {settings.pageXofY ? "Page 1 of 12" : "1"}
              </div>
            )}

            <div style={{ fontSize: `${settings.bodySize}pt` }}>
              <h1 className="font-bold" style={{ fontSize: `${settings.h1Size}pt`, color: settings.h1Color, marginBottom: `${settings.spacingAfter}pt` }}>
                Chapter 1: Introduction to Formatting
              </h1>
              <p style={{ marginBottom: `${settings.spacingAfter}pt` }}>
                This is how your body text will appear. Spacing, font size, and family update live. Indian universities typically require Times New Roman 12pt with 1.5 line spacing.
              </p>
              <h2 className="font-bold" style={{ fontSize: `${settings.h2Size}pt`, color: settings.h2Color, marginBottom: `${settings.spacingAfter}pt` }}>
                1.1 Section Heading
              </h2>
              <p style={{ marginBottom: `${settings.spacingAfter}pt` }}>
                Another paragraph of body text with the same styling throughout. Heading colors appear above.
              </p>
              <h3 className="font-bold" style={{ fontSize: `${settings.h3Size}pt`, color: settings.h3Color, marginBottom: `${settings.spacingAfter}pt` }}>
                1.1.1 Subsection
              </h3>
              <p style={{ marginBottom: `${settings.spacingAfter}pt` }}>
                Content under a subsection heading. Formatting rules apply consistently throughout.
              </p>
            </div>

            <div className="absolute bottom-6 left-0 right-0 text-center text-gray-300 text-xs select-none pointer-events-none">
              Preview is approximate — actual .docx may vary slightly
            </div>
          </div>
        </div>
      </ResizablePanel>
    </ResizablePanelGroup>
  );
}
