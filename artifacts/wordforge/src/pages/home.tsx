import { useState } from "react";
import { FormatterTab } from "@/components/formatter-tab";
import { NotebookTab } from "@/components/notebook-tab";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";

export default function Home() {
  return (
    <div className="flex flex-col h-screen max-h-screen bg-background text-foreground overflow-hidden">
      <header className="flex-none border-b border-border bg-card px-6 py-3 flex items-center justify-between z-10 shadow-sm">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-primary rounded-md flex items-center justify-center text-primary-foreground font-serif font-bold text-lg">
            W
          </div>
          <h1 className="font-semibold text-xl tracking-tight text-foreground">WordForge</h1>
        </div>
      </header>

      <main className="flex-1 overflow-hidden">
        <Tabs defaultValue="formatter" className="h-full flex flex-col">
          <div className="px-6 pt-4 bg-card/50 border-b border-border/50">
            <TabsList className="grid w-[400px] grid-cols-2">
              <TabsTrigger value="formatter">Formatter</TabsTrigger>
              <TabsTrigger value="notebook">Notebook Converter</TabsTrigger>
            </TabsList>
          </div>
          
          <div className="flex-1 overflow-hidden relative">
            <TabsContent value="formatter" className="h-full m-0 data-[state=active]:flex flex-col absolute inset-0">
              <FormatterTab />
            </TabsContent>
            
            <TabsContent value="notebook" className="h-full m-0 data-[state=active]:flex flex-col absolute inset-0 bg-muted/20">
              <NotebookTab />
            </TabsContent>
          </div>
        </Tabs>
      </main>
    </div>
  );
}