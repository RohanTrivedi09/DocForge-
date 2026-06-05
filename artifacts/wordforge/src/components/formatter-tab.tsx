import { useState, useRef, useEffect } from "react";
import { Download, Upload, Eye, FileText, Settings, AlignLeft, AlignCenter, AlignRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { ResizablePanelGroup, ResizablePanel, ResizableHandle } from "@/components/ui/resizable";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { toast } from "sonner";
import { useAiSuggest } from "@workspace/api-client-react";

export function FormatterTab() {
  const [file, setFile] = useState<File | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const suggest = useAiSuggest();
  
  // Track triggered suggestions to avoid spamming
  const triggeredHints = useRef<Set<string>>(new Set());

  // Formatter settings state
  const [settings, setSettings] = useState({
    fontFamily: "Times New Roman",
    bodySize: 12,
    bodyColor: "#000000",
    h1Size: 16,
    h1Color: "#000000",
    h2Size: 14,
    h2Color: "#000000",
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
    diffFirstPage: false,
    diffOddEven: false,
  });

  const updateSetting = (key: keyof typeof settings, value: any) => {
    setSettings(prev => ({ ...prev, [key]: value }));
    
    // Check triggers
    if (key === 'h1Size' && value > 20 && !triggeredHints.current.has('h1_large')) {
      triggeredHints.current.add('h1_large');
      triggerHint('h1_large', `H1 size: ${value}pt`);
    } else if (key === 'lineSpacing' && value === '2.0' && !triggeredHints.current.has('double_spacing')) {
      triggeredHints.current.add('double_spacing');
      triggerHint('double_spacing', `Line spacing: 2.0`);
    } else if (key.endsWith('Color') && typeof value === 'string') {
      // Basic luminance check (very rough)
      const hex = value.replace('#', '');
      const r = parseInt(hex.substr(0, 2), 16);
      const g = parseInt(hex.substr(2, 2), 16);
      const b = parseInt(hex.substr(4, 2), 16);
      const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
      
      if (luminance < 0.2 && !triggeredHints.current.has('dark_color')) {
        triggeredHints.current.add('dark_color');
        triggerHint('dark_color', `Color: ${value}`);
      }
    }
  };

  const triggerHint = (trigger: string, context: string = "") => {
    suggest.mutate({
      data: { trigger, context }
    }, {
      onSuccess: (data) => {
        if (data.suggestion) {
          toast(data.suggestion, {
            action: {
              label: "Accept",
              onClick: () => console.log("Hint accepted")
            },
            cancel: {
              label: "Dismiss",
              onClick: () => console.log("Hint dismissed")
            },
            duration: 10000,
          });
        }
      }
    });
  };

  const handleHeaderFocus = () => {
    if (!triggeredHints.current.has('header_focus')) {
      triggeredHints.current.add('header_focus');
      triggerHint('header_focus');
    }
  };

  const [isFormatting, setIsFormatting] = useState(false);

  const handleFormat = async () => {
    setIsFormatting(true);
    
    try {
      const formData = new FormData();
      if (file) {
        formData.append("file", file);
      }
      formData.append("settings", JSON.stringify(settings));

      const response = await fetch("/api/format-doc", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Formatting failed");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = file ? `Formatted_${file.name}` : "Formatted_Document.docx";
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      toast.success("Document formatted successfully");
    } catch (error) {
      console.error(error);
      toast.error("Failed to format document");
    } finally {
      setIsFormatting(false);
    }
  };

  return (
    <ResizablePanelGroup direction="horizontal" className="h-full flex-1 rounded-none border-none">
      {/* Controls Panel */}
      <ResizablePanel defaultSize={35} minSize={25} maxSize={50} className="bg-card flex flex-col">
        <div className="p-4 border-b border-border flex items-center justify-between bg-muted/30">
          <h2 className="font-semibold flex items-center gap-2">
            <Settings className="w-4 h-4 text-muted-foreground" />
            Formatting Options
          </h2>
        </div>
        
        <div className="flex-1 overflow-y-auto p-4 custom-scrollbar">
          <Accordion type="multiple" defaultValue={["typography", "layout", "header-footer"]} className="w-full">
            
            <AccordionItem value="upload" className="border-b-0 mb-4 bg-muted/20 p-4 rounded-lg border border-border">
              <div className="space-y-3">
                <Label className="text-sm font-semibold">Source Document (Optional)</Label>
                <div className="flex gap-2">
                  <Input 
                    type="file" 
                    accept=".docx" 
                    className="hidden" 
                    ref={fileInputRef}
                    onChange={(e) => {
                      if (e.target.files && e.target.files.length > 0) {
                        setFile(e.target.files[0]);
                      }
                    }}
                  />
                  <Button 
                    variant="outline" 
                    className="w-full justify-start text-muted-foreground font-normal"
                    onClick={() => fileInputRef.current?.click()}
                  >
                    <Upload className="w-4 h-4 mr-2" />
                    {file ? file.name : "Select .docx file..."}
                  </Button>
                  {file && (
                    <Button variant="ghost" size="icon" onClick={() => setFile(null)}>
                      &times;
                    </Button>
                  )}
                </div>
              </div>
            </AccordionItem>

            <AccordionItem value="typography">
              <AccordionTrigger className="hover:no-underline font-medium text-sm">Typography</AccordionTrigger>
              <AccordionContent className="pt-2 space-y-4">
                <div className="space-y-2">
                  <Label>Font Family</Label>
                  <select 
                    className="flex h-9 w-full items-center justify-between whitespace-nowrap rounded-md border border-input bg-transparent px-3 py-2 text-sm shadow-sm ring-offset-background placeholder:text-muted-foreground focus:outline-none focus:ring-1 focus:ring-ring disabled:cursor-not-allowed disabled:opacity-50"
                    value={settings.fontFamily}
                    onChange={(e) => updateSetting("fontFamily", e.target.value)}
                  >
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Arial">Arial</option>
                    <option value="Calibri">Calibri</option>
                    <option value="Georgia">Georgia</option>
                    <option value="Cambria">Cambria</option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label>Body Size (pt)</Label>
                    <Input type="number" value={settings.bodySize} onChange={(e) => updateSetting("bodySize", parseInt(e.target.value))} />
                  </div>
                  <div className="space-y-2">
                    <Label>Body Color</Label>
                    <div className="flex gap-2">
                      <Input type="color" className="w-10 p-1 h-9" value={settings.bodyColor} onChange={(e) => updateSetting("bodyColor", e.target.value)} />
                      <Input type="text" className="flex-1" value={settings.bodyColor} onChange={(e) => updateSetting("bodyColor", e.target.value)} />
                    </div>
                  </div>
                </div>

                <div className="space-y-3 pt-2">
                  <Label className="text-xs text-muted-foreground font-semibold uppercase tracking-wider">Headings</Label>
                  
                  {[1, 2, 3, 4].map((level) => {
                    const sizeKey = `h${level}Size` as keyof typeof settings;
                    const colorKey = `h${level}Color` as keyof typeof settings;
                    return (
                      <div key={level} className="grid grid-cols-[auto_1fr_1fr] gap-3 items-center">
                        <div className="w-8 font-medium">H{level}</div>
                        <Input type="number" placeholder="Size" value={settings[sizeKey] as number} onChange={(e) => updateSetting(sizeKey, parseInt(e.target.value))} className="h-8" />
                        <div className="flex gap-2">
                          <Input type="color" className="w-8 p-0 border-0 h-8 rounded overflow-hidden" value={settings[colorKey] as string} onChange={(e) => updateSetting(colorKey, e.target.value)} />
                        </div>
                      </div>
                    );
                  })}
                </div>

                <div className="space-y-2 pt-2">
                  <Label>Line Spacing</Label>
                  <select 
                    className="flex h-9 w-full items-center justify-between whitespace-nowrap rounded-md border border-input bg-transparent px-3 py-2 text-sm shadow-sm ring-offset-background placeholder:text-muted-foreground focus:outline-none focus:ring-1 focus:ring-ring disabled:cursor-not-allowed disabled:opacity-50"
                    value={settings.lineSpacing}
                    onChange={(e) => updateSetting("lineSpacing", e.target.value)}
                  >
                    <option value="1.0">Single (1.0)</option>
                    <option value="1.15">1.15</option>
                    <option value="1.5">1.5</option>
                    <option value="2.0">Double (2.0)</option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label>Spacing Before (pt)</Label>
                    <Input type="number" value={settings.spacingBefore} onChange={(e) => updateSetting("spacingBefore", parseInt(e.target.value))} />
                  </div>
                  <div className="space-y-2">
                    <Label>Spacing After (pt)</Label>
                    <Input type="number" value={settings.spacingAfter} onChange={(e) => updateSetting("spacingAfter", parseInt(e.target.value))} />
                  </div>
                </div>
              </AccordionContent>
            </AccordionItem>

            <AccordionItem value="layout">
              <AccordionTrigger className="hover:no-underline font-medium text-sm">Page Layout</AccordionTrigger>
              <AccordionContent className="pt-2 space-y-4">
                <Label className="text-xs text-muted-foreground font-semibold uppercase tracking-wider">Margins (cm)</Label>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <Label className="text-xs">Top</Label>
                    <Input type="number" step="0.1" value={settings.marginTop} onChange={(e) => updateSetting("marginTop", parseFloat(e.target.value))} />
                  </div>
                  <div className="space-y-2">
                    <Label className="text-xs">Bottom</Label>
                    <Input type="number" step="0.1" value={settings.marginBottom} onChange={(e) => updateSetting("marginBottom", parseFloat(e.target.value))} />
                  </div>
                  <div className="space-y-2">
                    <Label className="text-xs">Left</Label>
                    <Input type="number" step="0.1" value={settings.marginLeft} onChange={(e) => updateSetting("marginLeft", parseFloat(e.target.value))} />
                  </div>
                  <div className="space-y-2">
                    <Label className="text-xs">Right</Label>
                    <Input type="number" step="0.1" value={settings.marginRight} onChange={(e) => updateSetting("marginRight", parseFloat(e.target.value))} />
                  </div>
                </div>
              </AccordionContent>
            </AccordionItem>

            <AccordionItem value="header-footer">
              <AccordionTrigger className="hover:no-underline font-medium text-sm">Headers & Footers</AccordionTrigger>
              <AccordionContent className="pt-2 space-y-4">
                <div className="space-y-2">
                  <Label>Header Text</Label>
                  <Input 
                    placeholder="E.g., Chapter 1" 
                    value={settings.headerText} 
                    onChange={(e) => updateSetting("headerText", e.target.value)} 
                    onFocus={handleHeaderFocus}
                  />
                </div>
                <div className="space-y-2">
                  <Label>Footer Text</Label>
                  <Input 
                    placeholder="E.g., Confidential" 
                    value={settings.footerText} 
                    onChange={(e) => updateSetting("footerText", e.target.value)} 
                  />
                </div>

                <div className="space-y-4 pt-2">
                  <div className="flex items-center justify-between">
                    <Label htmlFor="page-numbers" className="flex-1 cursor-pointer">Page Numbers</Label>
                    <Switch id="page-numbers" checked={settings.pageNumbers} onCheckedChange={(c) => updateSetting("pageNumbers", c)} />
                  </div>
                  
                  {settings.pageNumbers && (
                    <div className="flex gap-2 p-1 bg-muted rounded-md w-max">
                      <Button variant={settings.pageNumberPos === 'left' ? 'secondary' : 'ghost'} size="sm" className="h-7 px-2" onClick={() => updateSetting("pageNumberPos", "left")}>
                        <AlignLeft className="w-4 h-4" />
                      </Button>
                      <Button variant={settings.pageNumberPos === 'center' ? 'secondary' : 'ghost'} size="sm" className="h-7 px-2" onClick={() => updateSetting("pageNumberPos", "center")}>
                        <AlignCenter className="w-4 h-4" />
                      </Button>
                      <Button variant={settings.pageNumberPos === 'right' ? 'secondary' : 'ghost'} size="sm" className="h-7 px-2" onClick={() => updateSetting("pageNumberPos", "right")}>
                        <AlignRight className="w-4 h-4" />
                      </Button>
                    </div>
                  )}

                  <div className="flex items-center justify-between">
                    <Label htmlFor="diff-first" className="flex-1 cursor-pointer text-sm">Different First Page</Label>
                    <Switch id="diff-first" checked={settings.diffFirstPage} onCheckedChange={(c) => updateSetting("diffFirstPage", c)} />
                  </div>
                  
                  <div className="flex items-center justify-between">
                    <Label htmlFor="diff-odd-even" className="flex-1 cursor-pointer text-sm">Different Odd & Even</Label>
                    <Switch id="diff-odd-even" checked={settings.diffOddEven} onCheckedChange={(c) => updateSetting("diffOddEven", c)} />
                  </div>
                </div>
              </AccordionContent>
            </AccordionItem>

          </Accordion>
        </div>

        <div className="p-4 border-t border-border bg-card">
          <Button className="w-full font-semibold gap-2" size="lg" onClick={handleFormat} disabled={isFormatting}>
            {isFormatting ? (
              <>Formatting...</>
            ) : (
              <>
                <Download className="w-4 h-4" />
                Format & Download .docx
              </>
            )}
          </Button>
        </div>
      </ResizablePanel>

      <ResizableHandle withHandle />

      {/* Preview Panel */}
      <ResizablePanel defaultSize={65} className="bg-muted/30 flex flex-col relative overflow-hidden">
        <div className="absolute top-4 left-4 z-10">
          <div className="bg-background/80 backdrop-blur-sm px-3 py-1.5 rounded-full text-xs font-medium border border-border shadow-sm flex items-center gap-2">
            <Eye className="w-3 h-3 text-muted-foreground" />
            Live Preview
          </div>
        </div>
        
        <div className="flex-1 overflow-auto p-8 flex items-start justify-center custom-scrollbar">
          {/* The Document Paper */}
          <div 
            className="bg-white shadow-xl transition-all duration-300 relative border border-gray-200"
            style={{
              width: "21cm", // A4 width
              minHeight: "29.7cm", // A4 height
              paddingTop: `${settings.marginTop}cm`,
              paddingBottom: `${settings.marginBottom}cm`,
              paddingLeft: `${settings.marginLeft}cm`,
              paddingRight: `${settings.marginRight}cm`,
              fontFamily: settings.fontFamily,
              color: settings.bodyColor,
              lineHeight: settings.lineSpacing,
            }}
          >
            {/* Header Area Indicator */}
            {settings.headerText && (
              <div className="absolute top-0 left-0 right-0 h-16 border-b border-gray-100 flex items-end px-12 pb-2 text-gray-400 text-sm">
                {settings.headerText}
              </div>
            )}
            
            {/* Page Number Indicator */}
            {settings.pageNumbers && (
              <div className={`absolute bottom-6 left-0 right-0 px-12 text-gray-400 text-sm flex ${
                settings.pageNumberPos === 'left' ? 'justify-start' : 
                settings.pageNumberPos === 'right' ? 'justify-end' : 'justify-center'
              }`}>
                1
              </div>
            )}

            <div style={{ fontSize: `${settings.bodySize}pt` }}>
              <h1 
                className="font-bold mb-4" 
                style={{ 
                  fontSize: `${settings.h1Size}pt`, 
                  color: settings.h1Color,
                  marginTop: `${settings.spacingBefore}pt`,
                  marginBottom: `${settings.spacingAfter}pt`
                }}
              >
                Chapter 1: Introduction to Formatting
              </h1>
              
              <p style={{ 
                marginTop: `${settings.spacingBefore}pt`,
                marginBottom: `${settings.spacingAfter}pt`
              }}>
                This is how your body text will appear. The spacing, font size, and family are all updated live to reflect your choices. 
                Indian universities often require specific strict formatting guidelines, usually Times New Roman 12pt with 1.5 line spacing.
              </p>

              <h2 
                className="font-bold mt-6 mb-3" 
                style={{ 
                  fontSize: `${settings.h2Size}pt`, 
                  color: settings.h2Color,
                  marginTop: `${settings.spacingBefore}pt`,
                  marginBottom: `${settings.spacingAfter}pt`
                }}
              >
                1.1 Section Heading
              </h2>
              
              <p style={{ 
                marginTop: `${settings.spacingBefore}pt`,
                marginBottom: `${settings.spacingAfter}pt`
              }}>
                Proper hierarchy is crucial for a thesis or long-form academic document. Notice how the margins constrain the text block, giving the page structure and readability.
              </p>

              <h3 
                className="font-bold mt-5 mb-2" 
                style={{ 
                  fontSize: `${settings.h3Size}pt`, 
                  color: settings.h3Color,
                  marginTop: `${settings.spacingBefore}pt`,
                  marginBottom: `${settings.spacingAfter}pt`
                }}
              >
                1.1.1 Subsection Details
              </h3>
              
              <p style={{ 
                marginTop: `${settings.spacingBefore}pt`,
                marginBottom: `${settings.spacingAfter}pt`
              }}>
                WordForge takes the pain out of manual formatting by applying these rules programmatically to your entire document at once, ensuring perfect consistency without the tedious clicking.
              </p>
            </div>
            
            {/* Margin visualizers (optional, subtle) */}
            <div className="absolute top-0 bottom-0 left-0 border-r border-blue-500/10 pointer-events-none" style={{ width: `${settings.marginLeft}cm` }} />
            <div className="absolute top-0 bottom-0 right-0 border-l border-blue-500/10 pointer-events-none" style={{ width: `${settings.marginRight}cm` }} />
            <div className="absolute left-0 right-0 top-0 border-b border-blue-500/10 pointer-events-none" style={{ height: `${settings.marginTop}cm` }} />
            <div className="absolute left-0 right-0 bottom-0 border-t border-blue-500/10 pointer-events-none" style={{ height: `${settings.marginBottom}cm` }} />
          </div>
        </div>
      </ResizablePanel>
    </ResizablePanelGroup>
  );
}