import React, { useState, useCallback, useRef } from 'react';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { GoogleGenAI } from '@google/genai';
import { UploadCloud, FileType, CheckCircle2, AlertCircle, Loader2, Download, Image as ImageIcon, Terminal } from 'lucide-react';
import { cn } from './lib/utils';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "dummy" });

interface ExtractedItem {
  id: string;
  filename: string;
  text: string;
  base64Data?: string;
  mimeType?: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
  result?: {
    produto: string;
    gramatura: string;
    quantidade: number | string;
  };
  error?: string;
}

export default function App() {
  const [isDragging, setIsDragging] = useState(false);
  const [items, setItems] = useState<ExtractedItem[]>([]);
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [currentFile, setCurrentFile] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [productDatabase, setProductDatabase] = useState<string[]>([]);
  const productFileInputRef = useRef<HTMLInputElement>(null);
  
  const [logs, setLogs] = useState<{id: string, time: Date, message: string, details?: string, type: string}[]>([]);
  const logsEndRef = useRef<HTMLDivElement>(null);

  const addLog = (message: string, type: 'info' | 'warning' | 'error' | 'success' = 'info', details?: string) => {
    setLogs(prev => [...prev, {
      id: Math.random().toString(36).substring(7),
      time: new Date(),
      message,
      details,
      type
    }]);
  };
  

  const processZip = async (file: File) => {
    try {
      setItems([]);
      setCurrentFile(file.name);
      const zip = await JSZip.loadAsync(file);
      
      const chatFile = Object.values(zip.files).find(f => !f.dir && f.name.endsWith('.txt'));
      if (!chatFile) {
        alert("Nenhum arquivo .txt encontrado no ZIP. Certifique-se de exportar a conversa do WhatsApp corretamente.");
        return;
      }

      const chatText = await chatFile.async('string');
      const lines = chatText.split('\n');
      
      const newItems: ExtractedItem[] = [];
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;
        
        // Verifica se a linha indica um anexo
        // Padrão 1 (iOS e outros): <anexado: file.jpg> ou <arquivo anexado: file.jpg>
        const matchIOS = line.match(/(.*?)<(?:arquivo )?anexado:\s*([^>]+?\.(?:jpg|jpeg|png|webp))>/i);
        // Padrão 2 (Android): file.jpg (arquivo anexado)
        const matchAndroid = line.match(/(.*?)([^:\s]+?\.(?:jpg|jpeg|png|webp))\s*\((?:arquivo )?anexado\)/i);
        
        let textBefore = "";
        let filename = "";

        if (matchIOS) {
          textBefore = matchIOS[1].trim();
          filename = matchIOS[2].trim();
        } else if (matchAndroid) {
          textBefore = matchAndroid[1].trim();
          filename = matchAndroid[2].trim();
        }
        
        if (filename) {
          let text = "";
          
          if (textBefore) {
            // O WhatsApp coloca ": " após o nome do remetente
            const parts = textBefore.split(': ');
            if (parts.length > 1) {
              text = parts.slice(1).join(': ').trim();
            } else {
              text = textBefore.trim();
            }
          }
          
          // Se a legenda não estiver na mesma linha, verifica a próxima linha (caso do Android)
          if (!text) {
             for (let j = i + 1; j < lines.length && j < i + 4; j++) {
                const nextLine = lines[j].trim();
                // Se começar com padrão de data de nova mensagem, para de buscar
                if (/^\[?\d{2}\/\d{2}\/\d{2,4}/.test(nextLine)) {
                  break; 
                }
                if (nextLine && !nextLine.toLowerCase().includes('anexado')) {
                  text += (text ? " " : "") + nextLine;
                }
             }
          }

          newItems.push({ 
            id: Math.random().toString(36).substring(7), 
            filename, 
            text: text || "Sem legenda", 
            status: 'pending' 
          });
        }
      }

      if (newItems.length === 0) {
        alert("Nenhuma imagem com legenda encontrada no arquivo de texto.");
        return;
      }

      for (const item of newItems) {
        const imageFile = Object.values(zip.files).find(f => !f.dir && !f.name.includes('__MACOSX') && f.name.endsWith(item.filename));
        if (imageFile) {
          const blob = await imageFile.async('blob');
          const base64Data = await new Promise<string>((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => {
              const res = reader.result as string;
              resolve(res.split(',')[1]);
            };
            reader.readAsDataURL(blob);
          });
          
          let mimeType = 'image/jpeg';
          if (item.filename.toLowerCase().endsWith('.png')) mimeType = 'image/png';
          else if (item.filename.toLowerCase().endsWith('.webp')) mimeType = 'image/webp';
          
          item.base64Data = base64Data;
          item.mimeType = mimeType;
        } else {
          item.status = 'error';
          item.error = 'Imagem não encontrada no ZIP';
        }
      }
      setItems(newItems);
    } catch (err) {
      console.error(err);
      alert("Erro ao processar o arquivo ZIP.");
    }
  };

  const onDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith('.zip')) {
        processZip(file);
      } else {
        alert("Por favor, envie um arquivo .zip");
      }
    }
  }, []);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processZip(e.target.files[0]);
    }
  };

  const handleProductFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0];
      try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const json = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
        const products = json.map(row => row[0]).filter(val => typeof val === 'string' && val.trim() !== '');
        
        setProductDatabase(products);
      } catch (err) {
        console.error("Erro ao ler planilha:", err);
        alert("Erro ao ler o arquivo de produtos.");
      }
    }
  };

  const processWithAI = async () => {
    if (!process.env.GEMINI_API_KEY) {
      alert("Chave da API Gemini não configurada!");
      return;
    }
    setProcessing(true);
    setProgress(0);
    let processedCount = 0;
    
    addLog(`=== Iniciando processamento de ${items.filter(i => i.status === 'pending' || i.status === 'error').length} itens ===`, 'info');
    
    // Process strictly sequentially based on original items snapshot
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.status === 'processing' || item.status === 'completed' || !item.base64Data) {
        processedCount++;
        setProgress(Math.round((processedCount / items.length) * 100));
        continue;
      }
      
      setItems(prev => prev.map(current => current.id === item.id ? { ...current, status: 'processing' } : current));
      const payloadSizeKb = Math.round((item.base64Data?.length || 0) / 1024);
      addLog(`[${item.filename}] Preparando análise. Base64: ${payloadSizeKb} KB`, 'info');

      try {
        const catalogContext = productDatabase.length > 0 
          ? `\n\nATENÇÃO: Você deve tentar encontrar uma correspondência exata para o produto nesta base de dados:\n[${productDatabase.slice(0, 500).join(', ')}]\nSe houver correspondência, o campo "produto" deve ser EXATAMENTE igual ao item da lista.\nSe NÃO encontrar NENHUMA correspondência nesta lista, identifique o produto pela foto/texto e coloque um ASTERISCO no início do nome (exemplo: "* Nome Identificado").` 
          : "";

        const prompt = `Analise esta imagem de prateleira/produto e a mensagem do operador: '${item.text}'.${catalogContext}
Extraia as informações e retorne ESTRITAMENTE um JSON no seguinte formato:
{
  "quantidade": 10,
  "produto": "Nome e Marca",
  "gramatura": "Tamanho/Peso"
}
Não inclua crases para blocos de código nem qualquer outro texto além do JSON puramente válido.`;

        let attempt = 0;
        let success = false;
        
        while (attempt < 3 && !success) {
          try {
            addLog(`[${item.filename}] Chamando Gemini API (T${attempt + 1}/3)...`, 'info');
            const response = await ai.models.generateContent({
              model: 'gemini-3-flash-preview',
              contents: { parts: [{ text: prompt }, { inlineData: { mimeType: item.mimeType!, data: item.base64Data } }] },
              config: { responseMimeType: 'application/json' }
            });

            const jsonText = response.text || "{}";
            addLog(`[${item.filename}] Resposta da IA recebida.`, 'success', jsonText);
            
            let parsedResult;
            try {
              parsedResult = JSON.parse(jsonText);
            } catch (e) {
              const cleanedText = jsonText.replace(/```json/g, '').replace(/```/g, '').trim();
              parsedResult = JSON.parse(cleanedText);
            }

            setItems(prev => prev.map(current => current.id === item.id ? { 
              ...current, 
              status: 'completed', 
              result: { produto: parsedResult.produto || "Desconhecido", gramatura: parsedResult.gramatura || "-", quantidade: parsedResult.quantidade || 0 }
            } : current));
            success = true;
          } catch (err: any) {
            attempt++;
            const errorMsg = err.message || "Failed to call the Gemini API";
            const stringifiedErr = JSON.stringify(err, Object.getOwnPropertyNames(err));
            addLog(`[${item.filename}] Erro na API (T${attempt}): ${errorMsg}`, 'warning', stringifiedErr);
            
            const isQuotaError = errorMsg.toLowerCase().includes('quota') || 
                                 errorMsg.includes('429') || 
                                 errorMsg.toLowerCase().includes('too many') || 
                                 errorMsg.toLowerCase().includes('exceeded');
            
            if (isQuotaError) {
              if (attempt < 3) {
                addLog(`[${item.filename}] Rate limit detectado. Aguardando 15s...`, 'warning');
                // Wait 15 seconds if it's a quota error to give more time to recover
                await new Promise(r => setTimeout(r, 15000));
                continue;
              } else {
                addLog(`[${item.filename}] Muitas falhas de Quota. Pausando script global.`, 'error');
                setItems(prev => prev.map(current => current.id === item.id ? { ...current, status: 'error', error: 'Rate limit/Quota atingido' } : current));
                // If it's explicitly a quota limit, it's better to pause everything
                setProcessing(false);
                return;
              }
            } else {
              if (attempt < 3) {
                 addLog(`[${item.filename}] Aguardando 3s antes de retentar...`, 'info');
                 await new Promise(r => setTimeout(r, 3000));
                 continue;
              }
              // Record the error for this item and let the loop continue to the next item
              addLog(`[${item.filename}] Item falhou após 3 tentativas.`, 'error');
              setItems(prev => prev.map(current => current.id === item.id ? { ...current, status: 'error', error: errorMsg } : current));
            }
          }
        }
        
      } catch (err: any) {
        addLog(`[${item.filename}] Falha catastrófica: ${err.message}`, 'error', err.stack);
        setItems(prev => prev.map(current => current.id === item.id ? { ...current, status: 'error', error: err.message || "Falha desconhecida" } : current));
      }
      
      processedCount++;
      setProgress(Math.round((processedCount / items.length) * 100));
      // Delay of 5 seconds to ensure we don't hit the 15 RPM free tier limit
      addLog(`[System] Aguardando 5s para rate limit geral...`, 'info');
      await new Promise(r => setTimeout(r, 5000));
    }
    setProcessing(false);
    addLog(`=== Processamento finalizado ===`, 'info');
  };

  const exportExcel = () => {
    const data = items.filter(i => i.status === 'completed' || i.status === 'error').map(item => ({
      Foto: item.filename,
      'Texto Original': item.text,
      Quantidade: item.result?.quantidade ?? '-',
      Produto: item.result?.produto ?? '-',
      Gramatura: item.result?.gramatura ?? '-',
      Status: item.status,
      Erro: item.error || ''
    }));

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Inventario");
    XLSX.writeFile(workbook, "inventario_estruturado.xlsx");
  };

  return (
    <div className="flex h-screen w-full overflow-hidden bg-slate-50 font-sans text-slate-900">
      <aside className="w-[300px] flex-shrink-0 bg-slate-900 text-slate-300 flex flex-col border-r border-slate-800">
        <div className="p-6 border-b border-slate-800">
          <div className="flex items-center gap-2 mb-1">
            <div className="w-3 h-3 rounded-full bg-emerald-500"></div>
            <span className="font-bold text-white tracking-tight uppercase text-sm">Vision Inventory AI</span>
          </div>
          <p className="text-xs text-slate-500 font-medium">Automated Logistics Script v2.4</p>
        </div>
        
        <nav className="flex-1 py-4">
          <div className="px-6 py-2 text-[10px] uppercase tracking-widest text-slate-500 font-bold mb-2">Modules</div>
          <a href="#" className="flex items-center gap-3 px-6 py-3 bg-slate-800 text-white">
            <FileType className="w-4 h-4 text-emerald-400" />
            <span className="text-sm">Ingestão & Parsing</span>
          </a>
          <a href="#" className="flex items-center gap-3 px-6 py-3 hover:bg-slate-800 transition-colors">
            <ImageIcon className="w-4 h-4" />
            <span className="text-sm">Processamento Vision</span>
          </a>
          <a href="#" className="flex items-center gap-3 px-6 py-3 hover:bg-slate-800 transition-colors">
            <Download className="w-4 h-4" />
            <span className="text-sm">Consolidação Pandas</span>
          </a>
        </nav>

        <div className="p-6 mt-auto border-t border-slate-800 space-y-4">
          <div className="bg-slate-800/50 rounded-lg p-3 border border-slate-700/50">
            <div className="text-[10px] text-slate-500 font-bold uppercase mb-2">Base de Produtos (Excel)</div>
            <div className="flex flex-col gap-2">
              {productDatabase.length > 0 ? (
                <div className="flex items-center justify-between">
                  <span className="text-xs text-slate-300">{productDatabase.length} itens</span>
                  <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                </div>
              ) : (
                <span className="text-xs text-slate-400">Nenhuma base importada (Opção)</span>
              )}
              <button 
                onClick={() => productFileInputRef.current?.click()}
                className="text-xs w-full py-1.5 bg-slate-700 hover:bg-slate-600 text-white rounded transition-colors"
              >
                {productDatabase.length > 0 ? "Atualizar Base" : "Importar Planilha (Col A)"}
              </button>
              <input 
                type="file" 
                accept=".xlsx,.xls,.csv" 
                className="hidden" 
                ref={productFileInputRef} 
                onChange={handleProductFileSelect} 
              />
            </div>
          </div>

          <div className="bg-slate-800/50 rounded-lg p-3 border border-slate-700/50">
            <div className="text-[10px] text-slate-500 font-bold uppercase mb-2">Engine Status</div>
            <div className="flex justify-between text-xs mb-1"><span>Gemini 3 Flash</span><span className="text-emerald-400">Active</span></div>
            <div className="flex justify-between text-xs"><span>Rate Limit</span><span>98%</span></div>
          </div>
        </div>
      </aside>

      <main className="flex-1 flex flex-col overflow-hidden">
        <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-8 flex-shrink-0">
          <h1 className="text-lg font-semibold text-slate-800">Arquitetura do Script de Inventário</h1>
          <div className="flex items-center gap-4">
            {currentFile && <span className="text-xs font-mono bg-slate-100 text-slate-800 font-medium px-2 py-1 rounded">PATH: /{currentFile}</span>}
            <button
              onClick={processWithAI}
              disabled={processing || items.length === 0 || items.every(i => i.status === 'completed')}
              className={cn(
                "text-sm font-medium px-4 py-1.5 rounded-md transition-shadow shadow-sm flex items-center gap-2",
                processing
                  ? "bg-indigo-100 text-indigo-400 cursor-not-allowed"
                  : items.every(i => i.status === 'completed') && items.length > 0
                    ? "bg-slate-100 text-slate-400 cursor-not-allowed"
                    : "bg-indigo-600 hover:bg-indigo-700 text-white"
              )}
            >
              {processing ? (
                <><Loader2 className="w-4 h-4 animate-spin" /> Executando Script...</>
              ) : items.every(i => i.status === 'completed') && items.length > 0 ? (
                <><CheckCircle2 className="w-4 h-4" /> Concluído</>
              ) : (
                "Executar Script"
              )}
            </button>
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-4 md:p-8">
          {!currentFile && items.length === 0 && (
            <div 
              className={cn(
                "border-2 border-dashed rounded-xl p-12 transition-all duration-200 ease-in-out flex flex-col items-center justify-center text-center bg-white min-h-[400px]",
                isDragging ? "bg-indigo-50 border-indigo-500 scale-[1.02]" : "border-slate-300 hover:border-slate-400"
              )}
              onDragOver={e => { e.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={onDrop}
              onClick={() => fileInputRef.current?.click()}
            >
              <input 
                type="file" 
                accept=".zip" 
                className="hidden" 
                ref={fileInputRef} 
                onChange={handleFileSelect} 
              />
              <div className="w-16 h-16 bg-slate-100 text-slate-400 rounded-full flex items-center justify-center mb-4">
                <UploadCloud className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-medium text-slate-700 mb-2">Importar Exportação do WhatsApp</h3>
              <p className="text-slate-500 max-w-md">
                Arraste e solte o arquivo <strong>.zip</strong> contendo o chat e as imagens, ou clique para selecionar.
              </p>
            </div>
          )}

          {items.length > 0 && (
            <div className="grid grid-cols-12 gap-6 h-full pb-8">
              <section className="col-span-12 xl:col-span-4 flex flex-col gap-6">
                <div className="bg-white rounded-xl border border-slate-200 shadow-sm flex-1 flex flex-col overflow-hidden min-h-[300px]">
                  <div className="p-4 border-b border-slate-100 bg-slate-50/50 flex justify-between items-center">
                    <h2 className="text-xs font-bold uppercase tracking-tight text-slate-500">1. Regex Ingestion ({items.length})</h2>
                    <span className="text-[10px] bg-indigo-100 text-indigo-700 font-bold px-2 rounded ml-2 whitespace-nowrap">REGEX ACTIVE</span>
                  </div>
                  <div className="flex-1 p-4 font-mono text-[11px] leading-relaxed overflow-y-auto text-slate-600 space-y-3">
                    {items.map((item, idx) => (
                      <div key={item.id} className={cn("p-2 rounded border-l-2", item.status === 'processing' ? "bg-indigo-50 border-indigo-400" : item.status === 'completed' ? "bg-emerald-50 border-emerald-400" : item.status === 'error' ? "bg-red-50 border-red-400" : "bg-slate-50 border-slate-300")}>
                        <span className="text-indigo-600 font-semibold">[Parsed]</span> {item.text}<br/>
                        <span className="text-emerald-600 text-[10px] mt-1 inline-block">&lt;anexado: {item.filename}&gt;</span>
                      </div>
                    ))}
                  </div>
                </div>
                
                {/* Console Log */}
                <div className="bg-slate-950 rounded-xl border border-slate-800 shadow-sm flex-1 flex flex-col overflow-hidden min-h-[250px] max-h-[400px]">
                  <div className="p-3 border-b border-slate-800 bg-slate-900/80 flex justify-between items-center">
                    <h2 className="text-xs font-bold uppercase tracking-tight text-slate-400 flex items-center gap-2">
                      <Terminal className="w-3.5 h-3.5" /> System Logs
                    </h2>
                    <button onClick={() => setLogs([])} className="text-[10px] text-slate-500 hover:text-slate-300 transition-colors uppercase font-bold tracking-wider">Clear</button>
                  </div>
                  <div className="flex-1 p-3 font-mono text-[11px] leading-relaxed overflow-y-auto space-y-2">
                    {logs.length === 0 ? (
                      <div className="text-slate-600 italic">Aguardando eventos do sistema...</div>
                    ) : (
                      logs.map(log => (
                        <div key={log.id} className="border-b border-slate-800/50 pb-2 last:border-0 last:pb-0">
                          <div className="flex gap-2 items-start">
                            <span className="text-slate-500 shrink-0">[{log.time.toLocaleTimeString()}]</span>
                            <span className={cn(
                              "break-words w-full",
                              log.type === 'error' ? 'text-red-400 font-medium' : 
                              log.type === 'warning' ? 'text-amber-400' : 
                              log.type === 'success' ? 'text-emerald-400' : 
                              'text-blue-300'
                            )}>{log.message}</span>
                          </div>
                          {log.details && (
                            <div className="mt-1.5 pl-[72px] text-slate-400 text-[10px] break-all max-h-[100px] overflow-y-auto whitespace-pre-wrap bg-slate-900/50 p-1.5 rounded">
                              {log.details}
                            </div>
                          )}
                        </div>
                      ))
                    )}
                    <div ref={logsEndRef} />
                  </div>
                </div>
              </section>

              <section className="col-span-12 xl:col-span-8 flex flex-col gap-6">
                {(processing || items.some(i => i.status === 'processing' || i.status === 'completed')) && (
                  <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-4 md:p-6 shrink-0">
                    <div className="flex flex-col sm:flex-row justify-between sm:items-start mb-4 md:mb-6 gap-2">
                      <div>
                        <h2 className="text-sm font-bold text-slate-800 mb-1">Extração Visual Gemini Flash</h2>
                        <p className="text-xs text-slate-500">Processando imagem {items.find(i => i.status === 'processing')?.filename || '...'} + Texto do operador</p>
                      </div>
                      <div className="flex gap-2 shrink-0">
                        <div className="flex flex-col items-start sm:items-end">
                          <span className="text-[10px] font-bold text-slate-400 uppercase">Progresso</span>
                          <span className="text-sm font-mono font-medium">{progress}%</span>
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-6">
                      <div className="aspect-video bg-slate-100 rounded-lg flex flex-col items-center justify-center border-2 border-dashed border-slate-300 relative overflow-hidden group">
                         {items.find(i => i.status === 'processing')?.base64Data ? (
                            <img src={`data:${items.find(i => i.status === 'processing')?.mimeType};base64,${items.find(i => i.status === 'processing')?.base64Data}`} className="absolute object-contain w-full h-full opacity-70 group-hover:opacity-100 transition-opacity" alt="Preview" />
                         ) : null}
                         {processing && <Loader2 className="w-8 h-8 text-slate-600 animate-spin mb-2 z-10" />}
                        <span className="text-[10px] font-bold tracking-wider text-slate-800 z-10 bg-white/90 px-3 py-1.5 rounded shadow-sm border border-slate-200 uppercase">
                          {processing ? "Analisando Imagem..." : "Aguardando"}
                        </span>
                      </div>
                      <div className="bg-slate-900 rounded-lg p-5 font-mono text-[12px] text-emerald-400 shadow-inner flex flex-col min-h-[140px] items-stretch">
                         <div className="text-slate-500 text-[10px] mb-3 font-bold uppercase tracking-wider">IA ACTIVITY</div>
                         {processing ? (
                           <div className="flex-1 flex items-center justify-start w-full">
                             <div className="animate-pulse space-y-1.5 w-full">
                                <div><span className="text-emerald-600">{"{"}</span></div>
                                <div className="pl-4 opacity-50 block w-full h-2 bg-emerald-800 rounded"></div>
                                <div className="pl-4 opacity-70 block w-3/4 h-2 bg-emerald-800 rounded"></div>
                                <div className="pl-4 opacity-60 block w-1/2 h-2 bg-emerald-800 rounded"></div>
                                <div><span className="text-emerald-600">{"}"}</span></div>
                             </div>
                           </div>
                         ) : (
                            <div className="flex-1 flex items-center justify-center text-slate-600 text-[10px] select-none text-center">
                              Inicie a execução para monitorar os retornos JSON da IA em tempo real.
                            </div>
                         )}
                      </div>
                    </div>
                  </div>
                )}
                
                <div className="bg-white rounded-xl border border-slate-200 shadow-sm flex-1 flex flex-col overflow-hidden min-h-[300px] xl:max-h-[calc(100vh-140px)]">
                  <div className="p-3 md:p-4 border-b border-slate-100 bg-slate-50/50 flex flex-wrap justify-between items-center gap-2">
                     <h2 className="text-xs font-bold uppercase tracking-tight text-slate-500">2. Consolidação Pandas</h2>
                     {items.some(i => i.status === 'completed') && (
                       <button
                         onClick={exportExcel}
                         disabled={processing}
                         className={cn("flex items-center justify-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-bold transition-colors", processing ? "bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed" : "border border-slate-300 bg-white hover:bg-slate-50 text-slate-700 shadow-sm")}
                       >
                         <Download className="w-3 h-3" /> EXPORT EXCEL
                       </button>
                     )}
                   </div>
                  <div className="overflow-x-auto overflow-y-auto flex-1 h-full">
                    <table className="w-full text-left text-xs border-collapse min-w-[600px]">
                      <thead className="sticky top-0 bg-slate-50 shadow-[0_1px_0_0_#f1f5f9] z-10">
                        <tr>
                          <th className="p-4 font-bold text-slate-500 uppercase tracking-tighter whitespace-nowrap">Status</th>
                          <th className="p-4 font-bold text-slate-500 uppercase tracking-tighter min-w-[200px]">Produto</th>
                          <th className="p-4 font-bold text-slate-500 uppercase tracking-tighter whitespace-nowrap">Gramatura</th>
                          <th className="p-4 font-bold text-slate-500 uppercase tracking-tighter whitespace-nowrap">Quantidade</th>
                          <th className="p-4 font-bold text-slate-500 uppercase tracking-tighter text-right whitespace-nowrap">Arquivo Ref</th>
                        </tr>
                      </thead>
                      <tbody className="text-slate-700">
                        {items.map((item) => (
                          <tr key={item.id} className="border-b border-slate-50 hover:bg-slate-50/80 transition-colors">
                            <td className="p-4">
                              {item.status === 'pending' && <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-slate-100 text-slate-600 text-[10px] font-bold"><div className="w-1.5 h-1.5 rounded-full bg-slate-400" /> PENDENTE</span>}
                              {item.status === 'processing' && <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-indigo-50 text-indigo-600 text-[10px] font-bold"><Loader2 className="w-2.5 h-2.5 animate-spin" /> ...</span>}
                              {item.status === 'completed' && <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-600 text-[10px] font-bold"><CheckCircle2 className="w-2.5 h-2.5" /> OK</span>}
                              {item.status === 'error' && <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-red-50 text-red-600 text-[10px] font-bold" title={item.error}><AlertCircle className="w-2.5 h-2.5" /> ERRO</span>}
                            </td>
                            <td className="p-4 font-medium">{item.result?.produto || <span className="text-slate-300">Aguardando IA...</span>}</td>
                            <td className="p-4 text-slate-500">{item.result?.gramatura || <span className="text-slate-300">-</span>}</td>
                            <td className="p-4 font-mono font-bold text-slate-900">{item.result?.quantidade ?? <span className="text-slate-300">-</span>}</td>
                            <td className="p-4 text-right text-slate-400 whitespace-nowrap">{item.filename}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </section>
            </div>
          )}
        </div>

        <footer className="h-12 bg-white border-t border-slate-200 flex items-center justify-between px-8 flex-shrink-0">
          <div className="flex items-center gap-6 text-[11px] font-medium text-slate-400">
            <span className="flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-emerald-500"></span> Script Ready</span>
            {items.length > 0 && <span>Items Processados: {items.filter(i => i.status === 'completed' || i.status === 'error').length} / {items.length}</span>}
            {items.filter(i => i.status === 'error').length > 0 && <span className="text-red-500">Erros: {items.filter(i => i.status === 'error').length}</span>}
          </div>
          <div className="flex items-center gap-2 text-[11px] font-semibold text-slate-500 hidden sm:flex">
            {items.some(i => i.status === 'completed') ? (
               <>
                 <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                 Excel Export ready: <button onClick={exportExcel} disabled={processing} className="text-indigo-600 hover:text-indigo-800 underline cursor-pointer">inventario_estruturado.xlsx</button>
               </>
            ) : (
               <span>Aguardando importação.</span>
            )}
          </div>
        </footer>
      </main>
    </div>
  );
}
