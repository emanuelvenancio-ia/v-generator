import React, { useState, useRef } from 'react';
import { 
  CheckCircle2, 
  Upload, 
  ChevronRight, 
  Shapes, 
  Settings, 
  FileText, 
  Table as TableIcon,
  ArrowRight,
  Database,
  ArrowLeftRight,
  Play,
  Download,
  ArrowLeft,
  AlertCircle,
  HelpCircle,
  XCircle
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { inspectFile, InspectionResult } from './lib/inspector';
import { parseExcel, getSheetData } from './lib/excel';
import { generateMergedFile, generateIndividualFiles } from './lib/merger';
import { jsPDF } from 'jspdf';
import JSZip from 'jszip';

// --- Components ---

const StepIndicator: React.FC<{ currentStep: number; totalSteps: number }> = ({ currentStep, totalSteps }) => {
  return (
    <div className="flex items-center gap-2 mb-8">
      {Array.from({ length: totalSteps }).map((_, i) => (
        <React.Fragment key={i}>
          <div className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold transition-all duration-500 ${
            i + 1 <= currentStep ? 'bg-blue-600 text-white shadow-lg shadow-blue-200' : 'bg-gray-100 text-gray-400'
          }`}>
            {i + 1 < currentStep ? <CheckCircle2 className="w-5 h-5" /> : i + 1}
          </div>
          {i < totalSteps - 1 && (
            <div className={`w-12 h-0.5 rounded-full transition-all duration-500 ${
              i + 1 < currentStep ? 'bg-blue-600' : 'bg-gray-100'
            }`} />
          )}
        </React.Fragment>
      ))}
    </div>
  );
};

const ModelPreview: React.FC<{ result: InspectionResult; selectedShapeId?: string }> = ({ result, selectedShapeId }) => {
  const pageSize = result.pageSize || { width: 9144000, height: 6858000 };
  const aspectRatio = pageSize.width / pageSize.height;
  
  const firstContainer = result.containers[0];
  if (!firstContainer) return null;

  return (
    <div className="w-full bg-white border border-gray-100 rounded-3xl p-8 shadow-sm mb-4">
      <div className="relative mx-auto bg-white border border-gray-100 shadow-xl overflow-hidden rounded-xl" 
           style={{ 
             aspectRatio: `${aspectRatio}`,
             maxWidth: '100%',
             height: 'auto'
           }}>
        
        {result.thumbnailUrl ? (
          <img 
            src={result.thumbnailUrl} 
            alt="Base Preview" 
            className={`absolute inset-0 w-full h-full object-contain transition-opacity duration-500 ${selectedShapeId ? 'opacity-60 grayscale-[0.2]' : 'opacity-100'}`}
            referrerPolicy="no-referrer"
          />
        ) : (
          <div className="absolute inset-0 bg-gray-50 flex items-center justify-center">
            <Shapes className="w-12 h-12 text-gray-200" />
          </div>
        )}

        <div className="absolute inset-0">
          <AnimatePresence>
            {firstContainer.shapes.map((shape) => {
              if (shape.x === undefined || shape.y === undefined || shape.width === undefined || shape.height === undefined) return null;
              
              const left = (shape.x / pageSize.width) * 100;
              const top = (shape.y / pageSize.height) * 100;
              const width = (shape.width / pageSize.width) * 100;
              const height = (shape.height / pageSize.height) * 100;
              const isSelected = selectedShapeId === shape.name;

              if (!isSelected) return null;

              return (
                <motion.div
                  key={shape.id}
                  initial={{ opacity: 0, scale: 0.8 }}
                  animate={{ opacity: 1, scale: 1 }}
                  exit={{ opacity: 0, scale: 0.8 }}
                  className="absolute border-2 border-blue-600 bg-blue-500/20 shadow-[0_0_20px_rgba(37,99,235,0.4)] z-10 flex items-center justify-center ring-2 ring-white ring-offset-0"
                  style={{
                    left: `${left}%`,
                    top: `${top}%`,
                    width: `${width}%`,
                    height: `${height}%`
                  }}
                >
                   <div className="bg-blue-600 text-white text-[10px] font-black px-2 py-1 rounded shadow-lg translate-y-[-100%] absolute top-0 flex items-center gap-1">
                     <CheckCircle2 className="w-3 h-3" />
                     {shape.name}
                   </div>
                </motion.div>
              );
            })}
          </AnimatePresence>
        </div>
      </div>
    </div>
  );
};

const LoadingOverlay: React.FC<{ message: string }> = ({ message }) => (
  <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="fixed inset-0 bg-white/80 backdrop-blur-sm z-[100] flex flex-col items-center justify-center p-6 text-center">
    <div className="w-20 h-20 relative mb-8">
      <motion.div 
        animate={{ rotate: 360 }} 
        transition={{ duration: 2, repeat: Infinity, ease: 'linear' }}
        className="absolute inset-0 border-4 border-gray-100 rounded-full"
      />
      <motion.div 
        animate={{ rotate: 360 }} 
        transition={{ duration: 1.5, repeat: Infinity, ease: 'linear' }}
        className="absolute inset-0 border-4 border-t-blue-600 border-r-transparent border-b-transparent border-l-transparent rounded-full"
      />
      <div className="absolute inset-0 flex items-center justify-center">
        <Settings className="w-8 h-8 text-blue-600 animate-pulse" />
      </div>
    </div>
    <h3 className="text-xl font-bold text-gray-900 mb-2">Processando...</h3>
    <p className="text-gray-500 font-medium max-w-xs">{message}</p>
    <div className="mt-8 flex gap-1">
      {[0, 1, 2].map(i => (
        <motion.div
          key={i}
          animate={{ scale: [1, 1.5, 1], opacity: [0.3, 1, 0.3] }}
          transition={{ duration: 1, repeat: Infinity, delay: i * 0.2 }}
          className="w-2 h-2 bg-blue-600 rounded-full"
        />
      ))}
    </div>
  </motion.div>
);

// --- Application ---

type Step = 'MODEL_UPLOAD' | 'EXCEL_UPLOAD' | 'SHEET_SELECT' | 'MAPPING' | 'SUMMARY' | 'DOWNLOAD_CHOICE' | 'FINISH';

export default function App() {
  const [step, setStep] = useState<Step>('MODEL_UPLOAD');
  const [loading, setLoading] = useState(false);
  const [loadingMessage, setLoadingMessage] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [showNoVarWarning, setShowNoVarWarning] = useState(false);
  const [sheetError, setSheetError] = useState(false);
  
  // Model state
  const [modelResult, setModelResult] = useState<InspectionResult | null>(null);
  const [modelFile, setModelFile] = useState<File | null>(null);
  
  // Excel state
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [excelSheets, setExcelSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [sheetData, setSheetData] = useState<{ columns: string[], rows: any[] } | null>(null);
  
  // Mapping state
  const [mappings, setMappings] = useState<Record<string, string>>({}); // column -> shapeName
  const [activeColumn, setActiveColumn] = useState<string | null>(null);
  const [previewShape, setPreviewShape] = useState<string | undefined>();

  // Range state
  const [mergeRangeType, setMergeRangeType] = useState<'ALL' | 'RANGE'>('ALL');
  const [rangeStart, setRangeStart] = useState<number>(1);
  const [rangeEnd, setRangeEnd] = useState<number>(1);

  // Filter state
  const [showRules, setShowRules] = useState(false);
  const [filterColumn, setFilterColumn] = useState<string | null>(null);
  const [selectedTags, setSelectedTags] = useState<string[]>([]);

  // Result state
  const [rowsToMerge, setRowsToMerge] = useState<any[]>([]);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) processFile(file);
    if (event.target) event.target.value = '';
  };

  const processFile = async (file: File) => {
    setLoading(true);
    setShowNoVarWarning(false);
    setSheetError(false);
    try {
      if (step === 'MODEL_UPLOAD') {
        setLoadingMessage('Inspecionando modelo e identificando variáveis...');
        const result = await inspectFile(file);
        
        // Validation: Check if there are any shapes with names
        const hasVariables = result.containers.some(c => c.shapes.length > 0);
        
        if (!hasVariables) {
          setShowNoVarWarning(true);
          setLoading(false);
          return;
        }

        setModelResult(result);
        setModelFile(file);
        setStep('EXCEL_UPLOAD');
      } else if (step === 'EXCEL_UPLOAD') {
        setLoadingMessage('Lendo base de dados Excel...');
        const data = await parseExcel(file);
        setExcelFile(file);
        setExcelSheets(data.sheets);
        setStep('SHEET_SELECT');
      }
    } catch (error) {
      console.error('File Processing Error:', error);
      alert('Erro ao processar arquivo. Verifique o formato.');
    } finally {
      setLoading(false);
      setLoadingMessage('');
      setIsDragging(false);
    }
  };

  const handleSheetSelect = async (sheetName: string) => {
    if (!excelFile) return;
    setLoading(true);
    setLoadingMessage('Validando conteúdo da planilha...');
    setSheetError(false);

    try {
      const buffer = await excelFile.arrayBuffer();
      const data = getSheetData(buffer, sheetName);
      
      if (data.columns.length === 0 || data.rows.length === 0) {
        setSheetError(true);
        setLoading(false);
        return;
      }
      
      setSelectedSheet(sheetName);
      setSheetData(data);
      setRangeStart(1);
      setRangeEnd(data.rows.length);
      setStep('MAPPING');
    } catch (e) {
      console.error('Sheet Error', e);
      alert('Erro ao ler planilha.');
    } finally {
      setLoading(false);
      setLoadingMessage('');
    }
  };

  const toggleMapping = (columnName: string, shapeName: string) => {
    setMappings(prev => {
      const next = { ...prev };
      // Check if this shape is already mapped to ANOTHER column, if so, free it from that column
      Object.keys(next).forEach(col => {
        if (next[col] === shapeName) delete next[col];
      });
      
      if (next[columnName] === shapeName) {
        delete next[columnName];
      } else {
        next[columnName] = shapeName;
      }
      return next;
    });
    // Don't auto-close, let user map multiple? No, UI is focused on one active.
    // setActiveColumn(null);
  };

  const handleStartMerge = async () => {
    if (!modelFile || !sheetData || Object.keys(mappings).length === 0) return;
    
    // 1. Respect the range first
    let finalRows = mergeRangeType === 'ALL' 
      ? sheetData.rows 
      : sheetData.rows.slice(Math.max(0, rangeStart - 1), Math.min(sheetData.rows.length, rangeEnd));

    // 2. Apply rules filter if active and configured
    if (showRules && filterColumn && selectedTags.length > 0) {
      finalRows = finalRows.filter(row => {
        const val = String(row[filterColumn] || '');
        return selectedTags.includes(val);
      });
    }

    if (finalRows.length === 0) {
      alert('O intervalo ou filtros selecionados não contêm dados válidos.');
      return;
    }

    setLoading(true);
    setLoadingMessage(`Preparando dados para exportação...`);
    try {
      setRowsToMerge(finalRows);
      setStep('DOWNLOAD_CHOICE');
    } catch (e) {
      console.error('Merge Error', e);
      alert('Erro ao preparar dados.');
    } finally {
      setLoading(false);
      setLoadingMessage('');
    }
  };

  const handleDownloadSingle = async () => {
    if (!modelFile || rowsToMerge.length === 0) return;
    setLoading(true);
    setLoadingMessage('Gerando arquivo único...');
    try {
      const resultBlob = await generateMergedFile(
        modelFile,
        { ...sheetData!, rows: rowsToMerge },
        mappings,
        modelResult?.fileType || 'pptx'
      );
      
      const url = URL.createObjectURL(resultBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Gerado_${modelFile.name}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setStep('FINISH');
    } catch (e) {
      console.error(e);
      alert('Erro ao gerar arquivo.');
    } finally {
      setLoading(false);
    }
  };

  const handleDownloadZip = async (format: 'png' | 'pdf') => {
    if (!modelFile || !modelResult || rowsToMerge.length === 0) return;
    setLoading(true);
    setLoadingMessage(`Gerando ZIP de ${format.toUpperCase()}s...`);
    
    try {
      const zip = new JSZip();
      
      // To generate PNG/PDF, we'll use a canvas-based approach if we have a thumbnail
      // Mapping coordinates are in EMUs.
      const pageSize = modelResult.pageSize || { width: 9144000, height: 6858000 };
      
      // Load thumbnail as image if available
      let templateImg: HTMLImageElement | null = null;
      if (modelResult.thumbnailUrl) {
        templateImg = await new Promise((resolve) => {
          const img = new Image();
          img.crossOrigin = 'anonymous';
          img.onload = () => resolve(img);
          img.onerror = () => resolve(null);
          img.src = modelResult.thumbnailUrl!;
        });
      }

      for (let i = 0; i < rowsToMerge.length; i++) {
        const row = rowsToMerge[i];
        
        // Find a filename
        const firstMappingCol = Object.keys(mappings)[0];
        const itemLabel = row[firstMappingCol] ? String(row[firstMappingCol]).replace(/[^a-z0-9]/gi, '_') : `Item_${i + 1}`;

        if (format === 'png') {
          // Create canvas
          const canvas = document.createElement('canvas');
          const scale = 2; // For better quality
          // 9144000 EMUs = 10 inches @ 96 DPI = 960px.
          // Let's use a reasonable base size.
          const exportWidth = (pageSize.width / 9144000) * 1280;
          const exportHeight = (pageSize.height / 6858000) * 960;
          
          canvas.width = exportWidth * scale;
          canvas.height = exportHeight * scale;
          const ctx = canvas.getContext('2d')!;
          
          // Draw background
          if (templateImg) {
            ctx.drawImage(templateImg, 0, 0, canvas.width, canvas.height);
          } else {
            ctx.fillStyle = 'white';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
          }

          // Overlay text
          ctx.scale(scale, scale);
          ctx.font = 'bold 24px Inter, sans-serif';
          ctx.fillStyle = '#1e293b';
          
          Object.entries(mappings).forEach(([col, shapeName]) => {
            const shape = modelResult.containers[0]?.shapes.find(s => s.name === shapeName);
            if (shape && shape.x !== undefined && shape.y !== undefined) {
              const x = (shape.x / pageSize.width) * exportWidth;
              const y = (shape.y / pageSize.height) * exportHeight;
              const width = (shape.width || 0) / pageSize.width * exportWidth;
              const height = (shape.height || 0) / pageSize.height * exportHeight;
              
              const val = String(row[col] || '');
              
              // Basic color logic for 'tipo_inscricao' as seen in merger.ts
              if (col === 'tipo_inscricao') {
                const raw = val.toUpperCase();
                if (raw.includes('MINI-CURSO')) ctx.fillStyle = '#0066CC';
                else if (raw.includes('AUTOR')) ctx.fillStyle = '#CC0000';
                else if (raw.includes('PRELECTOR')) ctx.fillStyle = '#660099';
                else if (raw.includes('EMPREENDEDOR')) ctx.fillStyle = '#FF8000';
                else ctx.fillStyle = '#00994C';
              } else {
                ctx.fillStyle = '#1e293b';
              }

              // Draw text centered in the shape area for best results
              ctx.textAlign = 'center';
              ctx.textBaseline = 'middle';
              ctx.fillText(val, x + width / 2, y + height / 2);
            }
          });

          const pngBlob = await new Promise<Blob>((resolve) => canvas.toBlob(b => resolve(b!), 'image/png'));
          zip.file(`${itemLabel}.png`, pngBlob);
        } else {
          // PDF
          const pdfW = (pageSize.width / 9144000) * 254; // mm
          const pdfH = (pageSize.height / 6858000) * 190.5; // mm
          const doc = new jsPDF({
            orientation: pdfW > pdfH ? 'landscape' : 'portrait',
            unit: 'mm',
            format: [pdfW, pdfH]
          });

          if (modelResult.thumbnailUrl) {
            doc.addImage(modelResult.thumbnailUrl, 'JPEG', 0, 0, pdfW, pdfH);
          }

          Object.entries(mappings).forEach(([col, shapeName]) => {
            const shape = modelResult.containers[0]?.shapes.find(s => s.name === shapeName);
            if (shape && shape.x !== undefined && shape.y !== undefined) {
              const x = (shape.x / pageSize.width) * pdfW;
              const y = (shape.y / pageSize.height) * pdfH;
              const width = (shape.width || 0) / pageSize.width * pdfW;
              const height = (shape.height || 0) / pageSize.height * pdfH;
              
              const val = String(row[col] || '');
              
              if (col === 'tipo_inscricao') {
                const raw = val.toUpperCase();
                if (raw.includes('MINI-CURSO')) doc.setTextColor(0, 102, 204);
                else if (raw.includes('AUTOR')) doc.setTextColor(204, 0, 0);
                else if (raw.includes('PRELECTOR')) doc.setTextColor(102, 0, 153);
                else if (raw.includes('EMPREENDEDOR')) doc.setTextColor(255, 128, 0);
                else doc.setTextColor(0, 153, 76);
              } else {
                doc.setTextColor(30, 41, 59);
              }

              doc.setFontSize(14);
              doc.text(val, x + width / 2, y + height / 2, { align: 'center', baseline: 'middle' });
            }
          });
          
          zip.file(`${itemLabel}.pdf`, doc.output('blob'));
        }
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Export_${format.toUpperCase()}_${new Date().getTime()}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setStep('FINISH');
    } catch (e) {
      console.error(e);
      alert('Erro ao gerar ZIP.');
    } finally {
      setLoading(false);
    }
  };

  const reset = () => {
    setStep('MODEL_UPLOAD');
    setShowNoVarWarning(false);
    setModelResult(null);
    setModelFile(null);
    setExcelFile(null);
    setExcelSheets([]);
    setSelectedSheet('');
    setSheetData(null);
    setMappings({});
  };

  const goBack = () => {
    if (step === 'EXCEL_UPLOAD') setStep('MODEL_UPLOAD');
    else if (step === 'SHEET_SELECT') setStep('EXCEL_UPLOAD');
    else if (step === 'MAPPING') setStep('SHEET_SELECT');
    else if (step === 'SUMMARY') setStep('MAPPING');
    else if (step === 'DOWNLOAD_CHOICE') setStep('SUMMARY');
  };

  return (
    <div className="min-h-screen bg-[#FDFDFF] text-gray-900 font-sans p-4 md:p-8">
      <div className="max-w-7xl mx-auto">
        
        {/* Header */}
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8">
          <div>
            <div className="flex items-center gap-3 mb-1">
              <div className="p-2 bg-blue-600 rounded-xl shadow-lg shadow-blue-100">
                <Database className="w-5 h-5 text-white" />
              </div>
              <h1 className="text-2xl font-black tracking-tight text-gray-900">DocShape <span className="text-blue-600 font-medium">Merge</span></h1>
            </div>
            <p className="text-gray-500 text-sm font-medium">Automatize a geração de documentos a partir de bases Excel.</p>
          </div>
          
          <div className="flex items-center gap-2">
            {step !== 'MODEL_UPLOAD' && step !== 'FINISH' && (
              <button 
                onClick={goBack} 
                className="px-4 py-2 text-xs font-bold text-gray-600 hover:text-blue-600 hover:bg-blue-50 transition-all rounded-lg flex items-center gap-2"
              >
                <ArrowLeft className="w-4 h-4" /> Anterior
              </button>
            )}
            <button onClick={reset} className="px-4 py-2 text-xs font-bold text-gray-400 hover:text-red-500 hover:bg-red-50 transition-all rounded-lg flex items-center gap-2">
              Reiniciar Processo
            </button>
          </div>
        </div>

        <StepIndicator 
          currentStep={
            step === 'MODEL_UPLOAD' ? 1 : 
            step === 'EXCEL_UPLOAD' ? 2 : 
            step === 'SHEET_SELECT' ? 3 : 
            step === 'MAPPING' ? 4 : 
            step === 'SUMMARY' ? 5 : 
            step === 'DOWNLOAD_CHOICE' ? 6 : 7
          } 
          totalSteps={7} 
        />

        {loading && loadingMessage && <LoadingOverlay message={loadingMessage} />}

        <AnimatePresence mode="wait">
          {step === 'MODEL_UPLOAD' && (
            <motion.div key="model_upload" initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -20 }} className="max-w-3xl mx-auto">
              <div className="text-center mb-8">
                <h2 className="text-xl font-bold text-gray-900 mb-1">1. Importe o Modelo Base</h2>
                <p className="text-gray-400 text-sm italic">Este arquivo funcionará como o seu template visual.</p>
              </div>

              {showNoVarWarning ? (
                <motion.div 
                  initial={{ opacity: 0, scale: 0.9 }} 
                  animate={{ opacity: 1, scale: 1 }}
                  className="bg-white border-2 border-red-100 rounded-[2rem] p-10 shadow-xl shadow-red-50/50 text-center"
                >
                  <div className="w-16 h-16 bg-red-50 rounded-2xl flex items-center justify-center mb-6 mx-auto">
                    <AlertCircle className="w-8 h-8 text-red-500" />
                  </div>
                  <h3 className="text-xl font-bold text-gray-900 mb-2">Não identificámos nenhuma variável no modelo</h3>
                  <p className="text-gray-500 text-sm mb-8">Não será possível editá-lo automaticamente se o Selection Pane estiver vazio ou sem nomes definidos.</p>
                  
                  <div className="grid grid-cols-1 gap-3 max-w-sm mx-auto mb-6">
                    <button 
                      onClick={() => { setShowNoVarWarning(false); fileInputRef.current?.click(); }}
                      className="w-full py-3 px-4 bg-blue-600 text-white rounded-xl font-bold text-sm flex items-center justify-center gap-2 hover:bg-blue-700 transition-all"
                    >
                      <FileText className="w-4 h-4" /> Selecionar outro modelo
                    </button>
                    <button 
                      onClick={reset}
                      className="w-full py-3 px-4 bg-gray-100 text-gray-600 rounded-xl font-bold text-sm flex items-center justify-center gap-2 hover:bg-gray-200 transition-all"
                    >
                      <XCircle className="w-4 h-4" /> Cancelar automação
                    </button>
                  </div>

                  <div className="pt-6 border-t border-gray-100">
                    <h4 className="text-xs font-black text-gray-400 uppercase tracking-widest mb-4 flex items-center justify-center gap-2">
                      <HelpCircle className="w-4 h-4" /> Tutorial de Preparação
                    </h4>
                    <div className="text-left text-sm text-gray-600 space-y-4 max-w-md mx-auto">
                      <div className="flex gap-3">
                        <div className="w-6 h-6 rounded-full bg-blue-50 text-blue-600 flex-shrink-0 flex items-center justify-center font-bold text-[10px]">1</div>
                        <p>No PowerPoint/Word, vá em <b>Página Inicial</b> {'>'} <b>Selecionar</b> {'>'} <b>Painel de Seleção</b>.</p>
                      </div>
                      <div className="flex gap-3">
                        <div className="w-6 h-6 rounded-full bg-blue-50 text-blue-600 flex-shrink-0 flex items-center justify-center font-bold text-[10px]">2</div>
                        <p>Dê nomes claros aos elementos que quer substituir (ex: <i>txtNome</i>, <i>txtData</i>).</p>
                      </div>
                      <div className="flex gap-3">
                        <div className="w-6 h-6 rounded-full bg-blue-50 text-blue-600 flex-shrink-0 flex items-center justify-center font-bold text-[10px]">3</div>
                        <p>Salve o arquivo e importe-o novamente aqui.</p>
                      </div>
                    </div>
                  </div>
                </motion.div>
              ) : (
                <div 
                  className={`p-16 border-4 border-dashed rounded-[2rem] transition-all flex flex-col items-center justify-center bg-white ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-100 hover:border-gray-200'}`}
                  onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                  onDragLeave={() => setIsDragging(false)}
                  onDrop={(e) => { e.preventDefault(); const file = e.dataTransfer.files[0]; if (file) processFile(file); }}
                >
                  <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" accept=".pptx,.docx" />
                  <div className="w-16 h-16 bg-blue-50 rounded-2xl flex items-center justify-center mb-6 shadow-inner">
                    {loading && !loadingMessage ? <Settings className="w-8 h-8 text-blue-600 animate-spin" /> : <Upload className="w-8 h-8 text-blue-600" />}
                  </div>
                  <h3 className="text-lg font-bold text-gray-900 mb-6">Arraste seu PPTX ou DOCX aqui</h3>
                  <button onClick={() => fileInputRef.current?.click()} disabled={loading} className="px-8 py-3 bg-gray-900 text-white rounded-xl font-bold hover:bg-blue-600 transition-all flex items-center gap-2 disabled:bg-gray-400">
                    <FileText className="w-4 h-4" /> Selecionar Arquivo
                  </button>
                </div>
              )}
            </motion.div>
          )}

          {step === 'EXCEL_UPLOAD' && (
            <motion.div key="excel_upload" initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -20 }} className="max-w-3xl mx-auto">
              <div className="text-center mb-8">
                <h2 className="text-xl font-bold text-gray-900 mb-1">2. Importe a Base de Dados</h2>
                <p className="text-gray-400 text-sm italic">O arquivo Excel com as informações para preenchimento.</p>
              </div>
              <div 
                className={`p-16 border-4 border-dashed rounded-[2rem] transition-all flex flex-col items-center justify-center bg-white ${isDragging ? 'border-green-500 bg-green-50' : 'border-gray-100 hover:border-gray-200'}`}
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={(e) => { e.preventDefault(); const file = e.dataTransfer.files[0]; if (file) processFile(file); }}
              >
                <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" accept=".xlsx" />
                <div className="w-16 h-16 bg-green-50 rounded-2xl flex items-center justify-center mb-6 shadow-inner">
                  {loading && !loadingMessage ? <Settings className="w-8 h-8 text-green-600 animate-spin" /> : <TableIcon className="w-8 h-8 text-green-600" />}
                </div>
                <h3 className="text-lg font-bold text-gray-900 mb-6">Arraste sua Planilha XLSX aqui</h3>
                <button onClick={() => fileInputRef.current?.click()} disabled={loading} className="px-8 py-3 bg-green-600 text-white rounded-xl font-bold hover:bg-green-700 transition-all flex items-center gap-2">
                  <Database className="w-4 h-4" /> Carregar Dados
                </button>
              </div>
            </motion.div>
          )}

          {step === 'SHEET_SELECT' && (
            <motion.div key="sheet_select" initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} className="max-w-xl mx-auto">
              {sheetError ? (
                <div className="bg-white border-2 border-amber-100 rounded-[2rem] p-10 shadow-xl shadow-amber-50/50 text-center">
                  <div className="w-16 h-16 bg-amber-50 rounded-2xl flex items-center justify-center mb-6 mx-auto">
                    <Database className="w-8 h-8 text-amber-500" />
                  </div>
                  <h3 className="text-xl font-bold text-gray-900 mb-2">Planilha sem dados detectados</h3>
                  <p className="text-gray-500 text-sm mb-8">A planilha selecionada não contém nenhuma tabela ou conteúdo. Escolha outra ou importe outro Excel.</p>
                  
                  <div className="grid grid-cols-1 gap-3 max-w-sm mx-auto">
                    {excelSheets.length > 1 && (
                      <button 
                        onClick={() => setSheetError(false)}
                        className="w-full py-3 px-4 bg-blue-600 text-white rounded-xl font-bold text-sm flex items-center justify-center gap-2 hover:bg-blue-700 transition-all"
                      >
                        <ChevronRight className="w-4 h-4 rotate-180" /> Selecionar outra planilha
                      </button>
                    )}
                    <button 
                      onClick={() => setStep('EXCEL_UPLOAD')}
                      className="w-full py-3 px-4 bg-gray-900 text-white rounded-xl font-bold text-sm flex items-center justify-center gap-2 hover:bg-blue-600 transition-all"
                    >
                      <TableIcon className="w-4 h-4" /> Selecionar outro Excel
                    </button>
                    <button 
                      onClick={reset}
                      className="w-full py-3 px-4 bg-gray-100 text-gray-600 rounded-xl font-bold text-sm flex items-center justify-center gap-2 hover:bg-gray-200 transition-all"
                    >
                      <XCircle className="w-4 h-4" /> Cancelar automação
                    </button>
                  </div>
                </div>
              ) : (
                <div className="bg-white border border-gray-100 rounded-[2rem] p-10 shadow-sm">
                  <h2 className="text-xl font-bold text-gray-900 mb-6 flex items-center gap-2">
                    <TableIcon className="w-6 h-6 text-green-600" /> Selecione a Planilha
                  </h2>
                  <div className="space-y-3">
                    {excelSheets.map((sheet) => (
                      <button key={sheet} onClick={() => handleSheetSelect(sheet)} className="w-full flex items-center justify-between p-4 border border-gray-100 rounded-xl hover:border-blue-200 hover:bg-blue-50 transition-all italic text-gray-600">
                        <span>{sheet}</span> <ChevronRight className="w-4 h-4 text-gray-300" />
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {step === 'MAPPING' && modelResult && sheetData && (
            <motion.div key="mapping" initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="flex flex-col gap-8">
              {/* Top Section: Side-by-Side Mapping Lists */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {/* Columns Selection */}
                <div className="bg-white border border-gray-100 rounded-[2rem] p-8 shadow-sm flex flex-col h-[500px]">
                  <h3 className="text-xs font-black uppercase tracking-widest text-blue-400 mb-6 flex items-center gap-2">
                    <TableIcon className="w-4 h-4" /> Colunas da Planilha
                  </h3>
                  <div className="space-y-2 overflow-y-auto pr-2 flex-grow scrollbar-thin scrollbar-thumb-gray-200">
                    {sheetData.columns.map((col) => (
                      <button
                        key={col}
                        onClick={() => setActiveColumn(activeColumn === col ? null : col)}
                        className={`w-full text-left p-4 rounded-xl border transition-all flex flex-col relative group ${
                          activeColumn === col 
                            ? 'border-blue-600 bg-blue-50 ring-2 ring-blue-100' 
                            : mappings[col] 
                              ? 'border-green-100 bg-green-50/20' 
                              : 'border-gray-50 hover:border-blue-200 hover:bg-gray-50'
                        }`}
                      >
                        <span className={`text-sm font-bold ${activeColumn === col ? 'text-blue-700' : 'text-gray-900'}`}>{col}</span>
                        {mappings[col] && (
                          <div className="mt-1 flex items-center gap-1.5">
                            <ArrowRight className="w-3 h-3 text-green-500" />
                            <span className="text-[10px] text-green-600 font-black uppercase tracking-wider">{mappings[col]}</span>
                          </div>
                        )}
                        {activeColumn === col && (
                          <motion.div layoutId="active-indicator" className="absolute right-4 top-1/2 -translate-y-1/2">
                            <div className="w-2 h-2 bg-blue-600 rounded-full animate-pulse" />
                          </motion.div>
                        )}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Selection Pane (Shapes) */}
                <div className="bg-white border border-gray-100 rounded-[2rem] p-8 shadow-sm flex flex-col h-[500px]">
                  <h3 className="text-xs font-black uppercase tracking-widest text-blue-400 mb-6 flex items-center gap-2">
                    <Shapes className="w-4 h-4" /> Painel de Seleção (Modelo)
                  </h3>
                  {!activeColumn ? (
                    <div className="flex flex-col items-center justify-center flex-grow text-center text-gray-300">
                      <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center mb-4 italic">!</div>
                      <p className="text-sm font-medium max-w-[200px]">Selecione uma coluna à esquerda para conectar</p>
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 overflow-y-auto pr-2 flex-grow scrollbar-thin scrollbar-thumb-gray-200 content-start">
                      {modelResult.containers[0]?.shapes.map((shape) => {
                        const isMappedToActive = mappings[activeColumn!] === shape.name;
                        const isMappedToOther = Object.entries(mappings).some(([col, name]) => col !== activeColumn && name === shape.name);

                        return (
                          <button
                            key={shape.id}
                            onMouseEnter={() => setPreviewShape(shape.name)}
                            onMouseLeave={() => setPreviewShape(undefined)}
                            onClick={() => toggleMapping(activeColumn!, shape.name)}
                            className={`p-4 border rounded-xl flex flex-col items-start text-left transition-all relative group ${
                              isMappedToActive 
                                ? 'border-blue-600 bg-blue-600 text-white shadow-lg shadow-blue-100' 
                                : isMappedToOther 
                                  ? 'border-gray-50 bg-gray-50 opacity-40 cursor-not-allowed' 
                                  : 'border-gray-100 hover:border-blue-400 hover:bg-blue-50/30'
                            }`}
                            disabled={isMappedToOther}
                          >
                            <span className={`text-[9px] font-black tracking-widest uppercase mb-1 ${isMappedToActive ? 'text-blue-100' : 'text-blue-400'}`}>
                              {shape.type}
                            </span>
                            <span className={`text-sm font-bold truncate w-full ${isMappedToActive ? 'text-white' : 'text-gray-900'}`}>
                              {shape.name}
                            </span>
                            {isMappedToActive && (
                              <div className="absolute top-2 right-2 translate-x-1 -translate-y-1">
                                <CheckCircle2 className="w-4 h-4 text-white" />
                              </div>
                            )}
                          </button>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>

              {/* Middle Section: Centralized Advance Button */}
              <div className="flex justify-center -my-2 relative z-10">
                <button 
                  onClick={() => setStep('SUMMARY')} 
                  disabled={Object.keys(mappings).length === 0} 
                  className="px-12 py-5 bg-blue-600 text-white rounded-2xl font-black text-lg flex items-center justify-center gap-3 hover:bg-blue-700 hover:scale-105 transition-all shadow-xl shadow-blue-200 disabled:bg-gray-100 disabled:text-gray-400 disabled:shadow-none disabled:scale-100"
                >
                  Avançar para Resumo <ArrowRight className="w-6 h-6" />
                </button>
              </div>

              {/* Bottom Section: Real-time Preview Full Width */}
              <div className="bg-white border border-gray-100 rounded-[2.5rem] p-10 shadow-sm">
                 <div className="flex items-center justify-between mb-6">
                    <h3 className="text-xs font-black uppercase tracking-widest text-gray-400">Pré-visualização em tempo real</h3>
                    <div className="flex gap-2">
                      <div className="w-2 h-2 bg-green-400 rounded-full animate-pulse" />
                      <div className="w-2 h-2 bg-gray-200 rounded-full" />
                      <div className="w-2 h-2 bg-gray-200 rounded-full" />
                    </div>
                 </div>
                 <ModelPreview 
                   result={modelResult} 
                   selectedShapeId={previewShape || (activeColumn ? mappings[activeColumn] : undefined)} 
                 />
              </div>
            </motion.div>
          )}

          {step === 'SUMMARY' && modelFile && sheetData && (
            <motion.div key="summary" initial={{ opacity: 0, scale: 0.98 }} animate={{ opacity: 1, scale: 1 }} className="max-w-2xl mx-auto">
               <div className="bg-white border border-gray-100 rounded-[2rem] p-10 shadow-sm text-center">
                  <Play className="w-12 h-12 text-blue-600 mx-auto mb-4" />
                  <h2 className="text-2xl font-black text-gray-900 mb-2">Confirmar Geração</h2>
                  <p className="text-gray-500 text-sm mb-8">Você confirma que deseja fazer as seguintes alterações no {modelFile.name}?</p>
                  <div className="bg-gray-50 rounded-2xl p-6 text-left mb-8 max-h-48 overflow-y-auto">
                     <div className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-4">Mapeamento Selecionado</div>
                     {Object.entries(mappings).map(([col, shape]) => (
                       <div key={col} className="flex items-center justify-between py-2 border-b border-gray-200/40 last:border-0 text-sm font-bold">
                         <span className="text-gray-900 italic">{col}</span>
                         <ArrowRight className="w-3 h-3 text-gray-300" />
                         <span className="text-blue-600">{shape}</span>
                       </div>
                     ))}
                  </div>

                  <div className="mb-8 text-left">
                    <div className="flex items-center justify-between mb-4">
                      <div className="text-xs font-bold text-gray-400 uppercase tracking-widest">Filtragem Inteligente</div>
                      <button 
                        onClick={() => setShowRules(!showRules)}
                        className={`text-[10px] font-black uppercase tracking-widest px-3 py-1 rounded-full transition-all ${showRules ? 'bg-amber-100 text-amber-600' : 'bg-gray-100 text-gray-400 hover:bg-blue-50 hover:text-blue-600'}`}
                      >
                        {showRules ? 'Desativar Regras' : 'Adicionar Regras'}
                      </button>
                    </div>

                    <AnimatePresence>
                      {showRules && (
                        <motion.div 
                          initial={{ opacity: 0, height: 0 }} 
                          animate={{ opacity: 1, height: 'auto' }} 
                          exit={{ opacity: 0, height: 0 }}
                          className="bg-gray-50/80 border border-gray-100 rounded-2xl p-6 mb-6 overflow-hidden"
                        >
                          <div className="mb-6">
                            <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-3">Selecione uma Coluna para Filtrar</label>
                            <div className="flex flex-wrap gap-2">
                              {sheetData.columns.map(col => (
                                <button
                                  key={col}
                                  onClick={() => {
                                    setFilterColumn(col);
                                    setSelectedTags([]);
                                  }}
                                  className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${filterColumn === col ? 'bg-blue-600 text-white shadow-md shadow-blue-100' : 'bg-white border border-gray-100 text-gray-500 hover:border-blue-200'}`}
                                >
                                  {col}
                                </button>
                              ))}
                            </div>
                          </div>

                          {filterColumn && (
                            <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}>
                              <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-3">Esta coluna contém os seguintes índices:</label>
                              <div className="flex flex-wrap gap-2 mb-2">
                                {Array.from(new Set(sheetData.rows.map(r => String(r[filterColumn] || ''))))
                                  .filter((val: string) => val.trim() !== '')
                                  .map(val => {
                                    const isSelected = selectedTags.includes(val);
                                    return (
                                      <button
                                        key={val}
                                        onClick={() => {
                                          setSelectedTags(prev => {
                                            if (prev.includes(val)) return prev.filter(t => t !== val);
                                            if (prev.length >= 2) return prev; // Limit to 2
                                            return [...prev, val];
                                          });
                                        }}
                                        className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${isSelected ? 'bg-amber-500 text-white shadow-md shadow-amber-100' : 'bg-white border border-gray-100 text-gray-500 hover:border-amber-200'}`}
                                      >
                                        {val}
                                      </button>
                                    );
                                  })}
                              </div>
                              <p className="text-[10px] text-gray-400 italic">Selecione no máximo 2 tags. O sistema aplicará uma interseção com o intervalo definido.</p>
                            </motion.div>
                          )}
                        </motion.div>
                      )}
                    </AnimatePresence>

                    <div className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-4">Quantidade de Arquivos</div>
                    <div className="grid grid-cols-2 gap-4 mb-4">
                      <button 
                        onClick={() => setMergeRangeType('ALL')}
                        className={`p-4 rounded-xl border text-sm font-bold transition-all ${mergeRangeType === 'ALL' ? 'border-blue-600 bg-blue-50 text-blue-700' : 'border-gray-100 text-gray-400 hover:border-gray-200'}`}
                      >
                        Todas as {sheetData.rows.length} linhas
                      </button>
                      <button 
                        onClick={() => setMergeRangeType('RANGE')}
                        className={`p-4 rounded-xl border text-sm font-bold transition-all ${mergeRangeType === 'RANGE' ? 'border-blue-600 bg-blue-50 text-blue-700' : 'border-gray-100 text-gray-400 hover:border-gray-200'}`}
                      >
                        Intervalo
                      </button>
                    </div>

                    {mergeRangeType === 'RANGE' && (
                      <motion.div initial={{ opacity: 0, height: 0 }} animate={{ opacity: 1, height: 'auto' }} className="flex items-center gap-4 bg-blue-50/50 p-4 rounded-xl border border-blue-100 mb-4">
                        <div className="flex-1">
                          <label className="block text-[10px] font-black text-blue-400 uppercase tracking-widest mb-1">Linha Inicial</label>
                          <input 
                            type="number" 
                            min="1" 
                            max={sheetData.rows.length}
                            value={rangeStart}
                            onChange={(e) => setRangeStart(Math.max(1, parseInt(e.target.value) || 1))}
                            className="w-full bg-white border border-blue-100 rounded-lg p-2 text-sm font-bold focus:ring-2 focus:ring-blue-200 outline-none"
                          />
                        </div>
                        <div className="flex-1">
                          <label className="block text-[10px] font-black text-blue-400 uppercase tracking-widest mb-1">Linha Final</label>
                          <input 
                            type="number" 
                            min={rangeStart} 
                            max={sheetData.rows.length}
                            value={rangeEnd}
                            onChange={(e) => setRangeEnd(Math.max(rangeStart, Math.min(sheetData.rows.length, parseInt(e.target.value) || sheetData.rows.length)))}
                            className="w-full bg-white border border-blue-100 rounded-lg p-2 text-sm font-bold focus:ring-2 focus:ring-blue-200 outline-none"
                          />
                        </div>
                      </motion.div>
                    )}
                  </div>

                  <div className="text-sm text-gray-400 mb-8 font-medium">
                    {mergeRangeType === 'ALL' ? (
                      <>Total de registros: <span className="text-blue-600 font-bold">{sheetData.rows.length}</span></>
                    ) : (
                      <>Serão gerados <span className="text-blue-600 font-bold">{Math.max(0, rangeEnd - rangeStart + 1)}</span> arquivos</>
                    )}
                  </div>
                  <div className="flex gap-3">
                     <button onClick={() => setStep('MAPPING')} className="flex-1 py-4 text-gray-400 font-bold">Voltar</button>
                     <button onClick={handleStartMerge} disabled={loading} className="flex-[2] py-4 bg-blue-600 text-white rounded-xl font-bold shadow-lg shadow-blue-100 flex items-center justify-center gap-2 disabled:bg-gray-300">
                        {loading ? <Settings className="w-5 h-5 animate-spin" /> : <><Play className="w-5 h-5" /> Gerar Arquivo</>}
                     </button>
                  </div>
               </div>
            </motion.div>
          )}

          {step === 'DOWNLOAD_CHOICE' && (
            <motion.div key="download_choice" initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="max-w-4xl mx-auto">
              <div className="text-center mb-12">
                <h2 className="text-3xl font-black text-gray-900 mb-2">Escolha o Método de Exportação</h2>
                <p className="text-gray-500 font-medium italic">Como você deseja receber os seus {rowsToMerge.length} arquivos?</p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                {/* Option 1: Single File */}
                <button 
                  onClick={handleDownloadSingle}
                  className="bg-white border-2 border-gray-100 rounded-[2.5rem] p-8 text-left hover:border-blue-600 hover:shadow-2xl hover:shadow-blue-50 transition-all group flex flex-col items-start gap-6"
                >
                  <div className="w-14 h-14 bg-blue-50 rounded-2xl flex items-center justify-center group-hover:bg-blue-600 transition-colors">
                    <FileText className="w-7 h-7 text-blue-600 group-hover:text-white" />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold text-gray-900 mb-2">Arquivo Único</h3>
                    <p className="text-gray-400 text-sm leading-relaxed mb-4">Um único arquivo {modelResult?.fileType.toUpperCase()} contendo todas as {rowsToMerge.length} gerações em sequência.</p>
                    <div className="text-[10px] font-black uppercase tracking-widest text-blue-600 flex items-center gap-1.5">
                      Melhor para Impressão <ChevronRight className="w-3 h-3" />
                    </div>
                  </div>
                </button>

                {/* Option 2: ZIP PNG */}
                <button 
                  onClick={() => handleDownloadZip('png')}
                  className="bg-white border-2 border-gray-100 rounded-[2.5rem] p-8 text-left hover:border-green-600 hover:shadow-2xl hover:shadow-green-50 transition-all group flex flex-col items-start gap-6"
                >
                  <div className="w-14 h-14 bg-green-50 rounded-2xl flex items-center justify-center group-hover:bg-green-600 transition-colors">
                    <Shapes className="w-7 h-7 text-green-600 group-hover:text-white" />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold text-gray-900 mb-2">Imagens (PNG)</h3>
                    <p className="text-gray-400 text-sm leading-relaxed mb-4">Um arquivo ZIP contendo cada certificado/documento individual como uma imagem de alta qualidade.</p>
                    <div className="text-[10px] font-black uppercase tracking-widest text-green-600 flex items-center gap-1.5">
                      Ideal para Redes Sociais <ChevronRight className="w-3 h-3" />
                    </div>
                  </div>
                </button>

                {/* Option 3: ZIP PDF */}
                <button 
                  onClick={() => handleDownloadZip('pdf')}
                  className="bg-white border-2 border-gray-100 rounded-[2.5rem] p-8 text-left hover:border-amber-600 hover:shadow-2xl hover:shadow-amber-50 transition-all group flex flex-col items-start gap-6"
                >
                  <div className="w-14 h-14 bg-amber-50 rounded-2xl flex items-center justify-center group-hover:bg-amber-600 transition-colors">
                    <Download className="w-7 h-7 text-amber-600 group-hover:text-white" />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold text-gray-900 mb-2">Digital (PDF)</h3>
                    <p className="text-gray-400 text-sm leading-relaxed mb-4">Um arquivo ZIP com cada documento exportado em formato PDF, pronto para envio por e-mail.</p>
                    <div className="text-[10px] font-black uppercase tracking-widest text-amber-600 flex items-center gap-1.5">
                      Padrão Corporativo <ChevronRight className="w-3 h-3" />
                    </div>
                  </div>
                </button>
              </div>
            </motion.div>
          )}

          {step === 'FINISH' && (
            <motion.div key="finish" initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} className="max-w-md mx-auto text-center">
               <div className="bg-white border border-gray-100 rounded-[2.5rem] p-12 shadow-sm">
                  <CheckCircle2 className="w-16 h-16 text-green-500 mx-auto mb-6" />
                  <h2 className="text-2xl font-black text-gray-900 mb-2">Sucesso!</h2>
                  <p className="text-gray-500 mb-8 font-medium">Os documentos foram gerados e o download já deve ter começado.</p>
                  <button onClick={reset} className="w-full py-4 bg-gray-900 text-white rounded-xl font-bold">Nova Automação</button>
               </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
}
