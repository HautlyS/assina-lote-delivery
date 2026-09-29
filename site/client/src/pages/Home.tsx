import { ChangeEvent, PointerEvent, useEffect, useMemo, useRef, useState } from "react";
import JSZip from "jszip";
import { renderAsync as renderDocxAsync } from "docx-preview";
import { PDFDocument } from "pdf-lib";
import {
  Archive,
  ArrowDownToLine,
  Check,
  ChevronRight,
  CircleAlert,
  FileArchive,
  FileCheck2,
  FileText,
  Fingerprint,
  ImagePlus,
  Layers3,
  LockKeyhole,
  MousePointer2,
  PenLine,
  Plus,
  RotateCcw,
  ShieldCheck,
  Sparkles,
  Trash2,
  Upload,
  WandSparkles,
} from "lucide-react";

type Placement = {
  x: number;
  y: number;
  width: number;
  opacity: number;
};

type Signature = {
  id: string;
  name: string;
  kind: "image" | "draw";
  dataUrl: string;
  placement: Placement;
};

type DocItem = {
  id: string;
  fileName: string;
  kind: "docx" | "pdf";
  bytes: Uint8Array;
  pages: number;
  selected: boolean;
  selectedPages: number[];
};

type PlacementException = {
  id: string;
  docId: string;
  page: number;
  signatureId: string;
  placement: Placement;
};

type AppliedPlacement = {
  page: number;
  signature: Signature;
  placement: Placement;
};

type RenderedPage = {
  page: number;
  top: number;
  left: number;
  width: number;
  height: number;
};

type ExportedFile = {
  name: string;
  bytes: Uint8Array;
  hash: string;
};

const DEFAULT_PLACEMENT: Placement = { x: 62, y: 78, width: 25, opacity: 100 };
const PAGE_WIDTH_PX = 490;
const PAGE_HEIGHT_PX = 660;

const uid = (prefix: string) => `${prefix}-${Math.random().toString(36).slice(2, 9)}`;

function bytesToDataUrl(bytes: Uint8Array, mime = "image/png") {
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...Array.from(bytes.subarray(i, i + chunk)));
  }
  return `data:${mime};base64,${btoa(binary)}`;
}

function dataUrlToBytes(dataUrl: string) {
  const base64 = dataUrl.split(",")[1] ?? "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export function escapeXml(value: string) {
  return value.replace(/[<>&'\"]/g, (char) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;" })[char] ?? char);
}

export function countLogicalPages(documentXml: string) {
  const matches = Array.from(documentXml.matchAll(/<w:br[^>]*w:type=["']page["'][^>]*\/?\s*>|<w:lastRenderedPageBreak\s*\/?\s*>/g));
  if (!matches.length) return 1;
  const last = matches[matches.length - 1];
  const tail = documentXml.slice((last.index ?? 0) + last[0].length);
  const trailingBreakIsEmpty = !/<w:(?:t|drawing|pict|object)\b[^>]*>[\s\S]*?\S/.test(tail);
  return Math.max(1, matches.length + 1 - (trailingBreakIsEmpty ? 1 : 0));
}

async function digestHex(bytes: Uint8Array) {
  const digest = await crypto.subtle.digest("SHA-256", bytes.buffer as ArrayBuffer);
  return Array.from(new Uint8Array(digest)).map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function stripBackground(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Não foi possível ler a imagem."));
    reader.onload = () => {
      const image = new Image();
      image.onerror = () => reject(new Error("Formato de imagem não reconhecido."));
      image.onload = () => {
        const scale = Math.min(1200 / image.width, 500 / image.height, 1);
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(1, Math.round(image.width * scale));
        canvas.height = Math.max(1, Math.round(image.height * scale));
        const ctx = canvas.getContext("2d", { willReadFrequently: true });
        if (!ctx) return reject(new Error("Canvas indisponível."));
        ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
        const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const { data, width, height } = pixels;
        const sample = (x: number, y: number) => {
          const i = (y * width + x) * 4;
          return [data[i], data[i + 1], data[i + 2]];
        };
        const samples = [sample(0, 0), sample(width - 1, 0), sample(0, height - 1), sample(width - 1, height - 1)];
        const bg = samples.reduce((acc, rgb) => acc.map((v, i) => v + rgb[i]), [0, 0, 0]).map((v) => v / samples.length);
        for (let i = 0; i < data.length; i += 4) {
          const distance = Math.hypot(data[i] - bg[0], data[i + 1] - bg[1], data[i + 2] - bg[2]);
          const light = (data[i] + data[i + 1] + data[i + 2]) / 3;
          if (data[i + 3] < 20 || distance < 52 || light > 245) {
            data[i + 3] = 0;
          } else if (distance < 90) {
            data[i + 3] = Math.round(data[i + 3] * ((distance - 52) / 38));
          }
        }
        ctx.putImageData(pixels, 0, 0);
        resolve(canvas.toDataURL("image/png"));
      };
      image.src = String(reader.result);
    };
    reader.readAsDataURL(file);
  });
}

function addImageToDocx(bytes: Uint8Array, placements: AppliedPlacement[]) {
  return JSZip.loadAsync(bytes).then(async (docZip) => {
    const documentFile = docZip.file("word/document.xml");
    if (!documentFile) throw new Error("O DOCX não contém word/document.xml.");
    let documentXml = await documentFile.async("text");
    let relsXml = await docZip.file("word/_rels/document.xml.rels")?.async("text");
    if (!relsXml) {
      relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }
    const ids = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g)).map((match) => Number(match[1]));
    let nextRel = Math.max(0, ...ids) + 1;
    let nextDrawing = 100;
    const paragraphs = Array.from(documentXml.matchAll(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g)).map((match) => match[0]);
    if (!paragraphs.length) throw new Error("Não foi possível localizar parágrafos no DOCX.");
    const pages = Array.from(new Set(placements.map((placement) => placement.page))).sort((a, b) => a - b);
    const drawingsByParagraph = new Map<number, string>();

    for (const page of pages) {
      let paragraphIndex = 0;
      if (page > 1) {
        let seenBreaks = 0;
        for (let i = 0; i < paragraphs.length; i += 1) {
          if (/<w:br[^>]*w:type=["']page["']|<w:lastRenderedPageBreak/.test(paragraphs[i])) {
            seenBreaks += 1;
            if (seenBreaks === page - 1) {
              paragraphIndex = Math.min(i + 1, paragraphs.length - 1);
              break;
            }
          }
        }
      }
      for (const [sigIndex, applied] of Array.from(placements.filter((placement) => placement.page === page).entries())) {
        const { signature, placement } = applied;
        const relId = `rId${nextRel++}`;
        const mediaName = `signature-${page}-${sigIndex + 1}-${signature.id}.png`;
        docZip.file(`word/media/${mediaName}`, dataUrlToBytes(signature.dataUrl));
        relsXml = relsXml.replace(
          "</Relationships>",
          `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${mediaName}"/></Relationships>`,
        );
        const pageX = Math.round((914400 * 8.27 * placement.x) / 100);
        const pageY = Math.round((914400 * 11.69 * placement.y) / 100);
        const cx = Math.round((914400 * 8.27 * placement.width) / 100);
        const cy = Math.round(cx * 0.375);
        const opacity = Math.max(0, Math.min(100, placement.opacity));
        const drawingId = nextDrawing++;
        const drawing = `<w:r><w:rPr/><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${pageX}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${pageY}</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${drawingId}" name="${escapeXml(signature.name)}" descr="Assinatura visual preparada pelo Assina Lote"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${drawingId}" name="${escapeXml(mediaName)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:alphaModFix amt="${opacity * 1000}"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
        drawingsByParagraph.set(paragraphIndex, `${drawingsByParagraph.get(paragraphIndex) ?? ""}${drawing}`);
      }
    }
    for (const [paragraphIndex, drawings] of Array.from(drawingsByParagraph.entries())) {
      const original = paragraphs[paragraphIndex];
      documentXml = documentXml.replace(original, original.replace("</w:p>", `${drawings}</w:p>`));
    }
    docZip.file("word/document.xml", documentXml);
    docZip.file("word/_rels/document.xml.rels", relsXml);
    return docZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  });
}

async function replaceTextInDocx(bytes: Uint8Array, search: string, replacement: string) {
  if (!search) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  const safeSearch = search.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const files = Object.values(zip.files).filter((entry) => !entry.dir && /^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/i.test(entry.name));
  for (const file of files) {
    const xml = await file.async("text");
    zip.file(file.name, xml.replace(new RegExp(safeSearch, "g"), escapeXml(replacement)));
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

async function addImagesToPdf(bytes: Uint8Array, placements: AppliedPlacement[]) {
  const pdf = await PDFDocument.load(bytes);
  const images = new Map<string, Awaited<ReturnType<typeof pdf.embedPng>>>();
  for (const page of pdf.getPages()) {
    const pageNumber = pdf.getPages().indexOf(page) + 1;
    const pagePlacements = placements.filter((item) => item.page === pageNumber);
    const { width, height } = page.getSize();
    for (const applied of pagePlacements) {
      let image = images.get(applied.signature.id);
      if (!image) { image = await pdf.embedPng(dataUrlToBytes(applied.signature.dataUrl)); images.set(applied.signature.id, image); }
      const drawWidth = (width * applied.placement.width) / 100;
      const drawHeight = drawWidth * image.height / image.width;
      page.drawImage(image, { x: (width * applied.placement.x) / 100, y: height - (height * applied.placement.y) / 100 - drawHeight, width: drawWidth, height: drawHeight, opacity: applied.placement.opacity / 100 });
    }
  }
  return pdf.save({ useObjectStreams: false });
}

function downloadBytes(bytes: Uint8Array, name: string, type: string) {
  const blob = new Blob([bytes.buffer as ArrayBuffer], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default function Home() {
  const [documents, setDocuments] = useState<DocItem[]>([]);
  const [signatures, setSignatures] = useState<Signature[]>([]);
  const [exceptions, setExceptions] = useState<PlacementException[]>([]);
  const [selectedSignatureIds, setSelectedSignatureIds] = useState<string[]>([]);
  const [activeSignatureId, setActiveSignatureId] = useState<string | null>(null);
  const [activePreviewDocId, setActivePreviewDocId] = useState<string | null>(null);
  const [activePreviewPage, setActivePreviewPage] = useState(1);
  const [renderedPageCount, setRenderedPageCount] = useState(0);
  const [renderedPages, setRenderedPages] = useState<RenderedPage[]>([]);
  const [previewRevision, setPreviewRevision] = useState(0);
  const [placementMode, setPlacementMode] = useState<"batch" | "exception">("batch");
  const [drawColor, setDrawColor] = useState("#142f3a");
  const [isDraggingPreview, setIsDraggingPreview] = useState(false);
  const [message, setMessage] = useState("Envie o ZIP dos DOCX para começar.");
  const [searchText, setSearchText] = useState("");
  const [replaceText, setReplaceText] = useState("");
  const [isExporting, setIsExporting] = useState(false);
  const [isApplyingReplacement, setIsApplyingReplacement] = useState(false);
  const [lastExport, setLastExport] = useState<ExportedFile[]>([]);
  const zipInputRef = useRef<HTMLInputElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const previewStylesRef = useRef<HTMLDivElement>(null);
  const renderRequestRef = useRef(0);
  const drawingRef = useRef(false);

  const activeSignature = useMemo(() => signatures.find((signature) => signature.id === activeSignatureId) ?? signatures[0], [activeSignatureId, signatures]);
  const activePreviewDoc = useMemo(() => documents.find((doc) => doc.id === activePreviewDocId) ?? documents[0], [activePreviewDocId, documents]);
  const selectedDocuments = documents.filter((doc) => doc.selected);
  const allPagesSelected = selectedDocuments.length > 0 && selectedDocuments.every((doc) => doc.selectedPages.length === doc.pages);

  useEffect(() => {
    let cancelled = false;
    const requestId = ++renderRequestRef.current;
    const container = previewRef.current;
    const styles = previewStylesRef.current;
    if (!container || !styles || !activePreviewDoc) {
      setRenderedPages([]);
      setRenderedPageCount(0);
      return;
    }
    container.innerHTML = "";
    styles.innerHTML = "";
    setRenderedPages([]);
    setRenderedPageCount(0);
    if (activePreviewDoc.kind === "pdf") {
      setMessage("PDF selecionado para exportação; a prévia visual do PDF permanece no arquivo original.");
      setRenderedPageCount(activePreviewDoc.pages);
      return () => { cancelled = true; };
    }
    const finish = (selector: string, fallbackCount?: number) => {
      if (cancelled || requestId !== renderRequestRef.current) return;
      const stage = container.parentElement;
      const stageRect = stage?.getBoundingClientRect();
      const pageElements = Array.from(container.querySelectorAll<HTMLElement>(selector));
      const metrics = pageElements.map((element, index) => {
        const rect = element.getBoundingClientRect();
        return { page: index + 1, top: rect.top - (stageRect?.top ?? 0) + (stage?.scrollTop ?? 0), left: rect.left - (stageRect?.left ?? 0), width: rect.width, height: rect.height };
      });
      const pageCount = metrics.length || fallbackCount || activePreviewDoc.pages;
      if (metrics.length && metrics.length !== activePreviewDoc.pages) {
        setDocuments((current) => current.map((doc) => {
          if (doc.id !== activePreviewDoc.id) return doc;
          const wasAllSelected = doc.selectedPages.length === doc.pages;
          const selectedPages = wasAllSelected ? Array.from({ length: pageCount }, (_, index) => index + 1) : doc.selectedPages.filter((page) => page <= pageCount);
          return { ...doc, pages: pageCount, selectedPages };
        }));
      }
      setRenderedPages(metrics);
      setRenderedPageCount(pageCount);
      setActivePreviewPage((current) => Math.min(current, pageCount));
    };
    void (async () => {
      try {
        await renderDocxAsync(activePreviewDoc.bytes, container, styles, {
          className: "docx",
          inWrapper: true,
          hideWrapperOnPrint: false,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          ignoreLastRenderedPageBreak: false,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
          renderComments: false,
          renderChanges: false,
          renderAltChunks: true,
          useBase64URL: true,
          experimental: true,
          debug: false,
        });
        await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
        setMessage("DOCX renderizado com layout, fontes, imagens e paginação do documento.");
        finish(".docx", activePreviewDoc.pages);
      } catch (error) {
        if (!cancelled) setMessage(error instanceof Error ? `Falha no renderizador DOCX: ${error.message}` : "Falha no renderizador DOCX.");
      }
    })();
    return () => { cancelled = true; };
  }, [activePreviewDoc, previewRevision]);

  const parseZip = async (file: File) => {
    setMessage("Lendo ZIP e identificando documentos...");
    const sourceZip = await JSZip.loadAsync(file);
    const entries = Object.values(sourceZip.files).filter((entry) => !entry.dir && /\.(docx|pdf)$/i.test(entry.name));
    if (!entries.length) throw new Error("Não encontrei arquivos .DOCX ou .PDF dentro deste ZIP.");
    const nextDocs: DocItem[] = [];
    for (const entry of entries) {
      const bytes = await entry.async("uint8array");
      let pages = 1;
      const kind: "docx" | "pdf" = entry.name.toLowerCase().endsWith(".pdf") ? "pdf" : "docx";
      try {
        if (kind === "pdf") pages = (await PDFDocument.load(bytes)).getPageCount();
        else {
          const innerZip = await JSZip.loadAsync(bytes);
          const xml = await innerZip.file("word/document.xml")?.async("text");
          if (xml) pages = countLogicalPages(xml);
        }
      } catch {
        pages = 1;
      }
      nextDocs.push({ id: uid("doc"), fileName: entry.name.split("/").pop() ?? entry.name, kind, bytes, pages, selected: true, selectedPages: Array.from({ length: pages }, (_, i) => i + 1) });
    }
    setDocuments(nextDocs);
    setActivePreviewDocId(nextDocs[0]?.id ?? null);
    setActivePreviewPage(1);
    setMessage(`${nextDocs.length} documentos carregados. Selecione páginas e assinaturas para aplicar.`);
  };

  const onZipChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await parseZip(file);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Falha ao abrir o ZIP.");
    } finally {
      event.target.value = "";
    }
  };

  const onImageChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setMessage("Tratando a imagem e removendo o fundo...");
      const dataUrl = await stripBackground(file);
      const signature: Signature = { id: uid("sig"), name: file.name.replace(/\.[^.]+$/, ""), kind: "image", dataUrl, placement: { ...DEFAULT_PLACEMENT } };
      setSignatures((current) => [...current, signature]);
      setSelectedSignatureIds((current) => [...current, signature.id]);
      setActiveSignatureId(signature.id);
      setMessage("Assinatura adicionada com fundo transparente. Ajuste o placement na prévia.");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Falha ao processar a assinatura.");
    } finally {
      event.target.value = "";
    }
  };

  const startDraw = (event: PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    event.preventDefault();
    const rect = canvas.getBoundingClientRect();
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.strokeStyle = drawColor;
    ctx.lineWidth = 4;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    ctx.beginPath();
    ctx.moveTo(((event.clientX - rect.left) / rect.width) * canvas.width, ((event.clientY - rect.top) / rect.height) * canvas.height);
    drawingRef.current = true;
    canvas.setPointerCapture(event.pointerId);
  };

  const moveDraw = (event: PointerEvent<HTMLCanvasElement>) => {
    if (!drawingRef.current) return;
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext("2d");
    if (!canvas || !ctx) return;
    event.preventDefault();
    const rect = canvas.getBoundingClientRect();
    ctx.lineTo(((event.clientX - rect.left) / rect.width) * canvas.width, ((event.clientY - rect.top) / rect.height) * canvas.height);
    ctx.stroke();
  };

  const finishDraw = (event?: PointerEvent<HTMLCanvasElement>) => {
    drawingRef.current = false;
    const canvas = canvasRef.current;
    if (canvas && event && canvas.hasPointerCapture(event.pointerId)) canvas.releasePointerCapture(event.pointerId);
  };

  const clearDraw = () => {
    const canvas = canvasRef.current;
    const ctx = canvas?.getContext("2d");
    if (canvas && ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
  };

  const saveDrawnSignature = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const hasInk = canvas.getContext("2d")?.getImageData(0, 0, canvas.width, canvas.height).data.some((value, index) => index % 4 === 3 && value > 0);
    if (!hasInk) {
      setMessage("Desenhe uma assinatura antes de salvar.");
      return;
    }
    const signature: Signature = { id: uid("sig"), name: "Assinatura desenhada", kind: "draw", dataUrl: canvas.toDataURL("image/png"), placement: { ...DEFAULT_PLACEMENT } };
    setSignatures((current) => [...current, signature]);
    setSelectedSignatureIds((current) => [...current, signature.id]);
    setActiveSignatureId(signature.id);
    clearDraw();
    setMessage("Assinatura desenhada adicionada ao lote.");
  };

  const getException = (docId: string | null | undefined, page: number, signatureId: string | undefined) => exceptions.find((exception) => exception.docId === docId && exception.page === page && exception.signatureId === signatureId);
  const getEffectivePlacement = (docId: string | null | undefined, page: number, signature: Signature) => getException(docId, page, signature.id)?.placement ?? signature.placement;

  const updatePlacement = (field: keyof Placement, value: number) => {
    if (!activeSignature) return;
    if (placementMode === "exception" && activePreviewDoc) {
      setExceptions((current) => {
        const existing = current.find((exception) => exception.docId === activePreviewDoc.id && exception.page === activePreviewPage && exception.signatureId === activeSignature.id);
        if (existing) return current.map((exception) => exception.id === existing.id ? { ...exception, placement: { ...exception.placement, [field]: value } } : exception);
        return [...current, { id: uid("exception"), docId: activePreviewDoc.id, page: activePreviewPage, signatureId: activeSignature.id, placement: { ...activeSignature.placement, [field]: value } }];
      });
      return;
    }
    setSignatures((current) => current.map((signature) => signature.id === activeSignature.id ? { ...signature, placement: { ...signature.placement, [field]: value } } : signature));
  };

  const createPageException = () => {
    if (!activeSignature || !activePreviewDoc) return;
    if (!getException(activePreviewDoc.id, activePreviewPage, activeSignature.id)) {
      setExceptions((current) => [...current, { id: uid("exception"), docId: activePreviewDoc.id, page: activePreviewPage, signatureId: activeSignature.id, placement: { ...activeSignature.placement } }]);
    }
    setPlacementMode("exception");
    setMessage(`Exceção criada para a página ${activePreviewPage} de ${activePreviewDoc.fileName}.`);
  };

  const removePageException = () => {
    if (!activeSignature || !activePreviewDoc) return;
    setExceptions((current) => current.filter((exception) => !(exception.docId === activePreviewDoc.id && exception.page === activePreviewPage && exception.signatureId === activeSignature.id)));
    setPlacementMode("batch");
    setMessage("A página voltou a seguir o placement em lote.");
  };

  const startPreviewDrag = (event: PointerEvent<HTMLDivElement>) => {
    if (!activeSignature) return;
    event.currentTarget.setPointerCapture(event.pointerId);
    setIsDraggingPreview(true);
  };

  const movePreviewDrag = (event: PointerEvent<HTMLDivElement>) => {
    if (!isDraggingPreview || !activeSignature) return;
    const page = event.currentTarget.parentElement;
    if (!page) return;
    const rect = page.getBoundingClientRect();
    const x = Math.max(8, Math.min(88, ((event.clientX - rect.left) / rect.width) * 100));
    const y = Math.max(8, Math.min(92, ((event.clientY - rect.top) / rect.height) * 100));
    updatePlacement("x", x);
    updatePlacement("y", y);
  };

  const applyPlacementToSelected = () => {
    if (!activeSignature) return;
    setSignatures((current) => current.map((signature) => selectedSignatureIds.includes(signature.id) ? { ...signature, placement: { ...activeSignature.placement } } : signature));
    setMessage("Placement replicado para as assinaturas selecionadas.");
  };

  const toggleDocument = (id: string) => setDocuments((current) => current.map((doc) => doc.id === id ? { ...doc, selected: !doc.selected } : doc));
  const togglePage = (id: string, page: number) => setDocuments((current) => current.map((doc) => {
    if (doc.id !== id) return doc;
    const has = doc.selectedPages.includes(page);
    return { ...doc, selectedPages: has ? doc.selectedPages.filter((item) => item !== page) : [...doc.selectedPages, page].sort((a, b) => a - b) };
  }));
  const toggleAllPages = () => setDocuments((current) => current.map((doc) => ({ ...doc, selectedPages: allPagesSelected ? [] : Array.from({ length: doc.pages }, (_, i) => i + 1) })));
  const toggleSignature = (id: string) => setSelectedSignatureIds((current) => current.includes(id) ? current.filter((item) => item !== id) : [...current, id]);
  const removeSignature = (id: string) => {
    setSignatures((current) => current.filter((signature) => signature.id !== id));
    setSelectedSignatureIds((current) => current.filter((item) => item !== id));
    if (activeSignatureId === id) setActiveSignatureId(null);
  };

  const applyTextReplacement = async () => {
    if (!searchText) return setMessage("Informe o texto que deseja localizar.");
    const targets = documents.filter((doc) => doc.selected && doc.kind === "docx");
    if (!targets.length) return setMessage("Selecione pelo menos um DOCX para aplicar a substituição.");
    setIsApplyingReplacement(true);
    try {
      const replacements = new Map<string, Uint8Array>();
      for (const doc of targets) replacements.set(doc.id, await replaceTextInDocx(doc.bytes, searchText, replaceText));
      setDocuments((current) => current.map((doc) => replacements.has(doc.id) ? { ...doc, bytes: replacements.get(doc.id)! } : doc));
      setPreviewRevision((current) => current + 1);
      setMessage(`${targets.length} DOCX atualizado(s); a prévia foi atualizada.`);
    } catch (error) {
      setMessage(error instanceof Error ? `Falha na substituição: ${error.message}` : "Falha na substituição.");
    } finally { setIsApplyingReplacement(false); }
  };

  const exportZip = async () => {
    if (!selectedDocuments.length) return setMessage("Selecione pelo menos um documento.");
    const chosenSignatures = signatures.filter((signature) => selectedSignatureIds.includes(signature.id));
    if (!chosenSignatures.length) return setMessage("Selecione pelo menos uma assinatura.");
    setIsExporting(true);
    setMessage("Aplicando assinaturas e calculando hashes de integridade...");
    try {
      const outputZip = new JSZip();
      const exported: ExportedFile[] = [];
      for (const doc of selectedDocuments) {
        const pages = doc.selectedPages.length ? doc.selectedPages : [1];
        const placements = pages.flatMap((page) => chosenSignatures.map((signature) => ({ page, signature, placement: getEffectivePlacement(doc.id, page, signature) })));
        let signedBytes = doc.bytes;
        if (doc.kind === "docx") {
          signedBytes = await replaceTextInDocx(signedBytes, searchText, replaceText);
          signedBytes = await addImageToDocx(signedBytes, placements);
        } else if (chosenSignatures.length) {
          signedBytes = await addImagesToPdf(signedBytes, placements);
        }
        const hash = await digestHex(signedBytes);
        outputZip.file(`assinados/${doc.fileName}`, signedBytes);
        exported.push({ name: doc.fileName, bytes: signedBytes, hash });
      }
      if (exported.length === 1) {
        const single = exported[0];
        downloadBytes(single.bytes, single.name, single.name.toLowerCase().endsWith(".pdf") ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        setLastExport(exported);
        setMessage("1 documento exportado diretamente, sem ZIP.");
        return;
      }
      const manifest = {
        produto: "Assina Lote",
        tipo: "integridade-de-lote-e-assinatura-visual",
        assinaturaQualificada: false,
        observacao: "Este manifesto comprova a integridade dos arquivos exportados por SHA-256. A imagem inserida é uma assinatura eletrônica visual e não substitui assinatura qualificada ICP-Brasil ou certificado digital emitido por autoridade certificadora.",
        criadoEm: new Date().toISOString(),
        documentos: exported.map(({ name, hash }) => ({ arquivo: name, sha256: hash, paginasSelecionadas: documents.find((doc) => doc.fileName === name)?.selectedPages ?? [1] })),
        assinaturas: chosenSignatures.map((signature) => ({ nome: signature.name, origem: signature.kind === "draw" ? "desenho no navegador" : "imagem tratada com remoção automática de fundo", placement: signature.placement })),
        excecoes: exceptions.map((exception) => ({ documento: documents.find((doc) => doc.id === exception.docId)?.fileName, pagina: exception.page, assinatura: signatures.find((signature) => signature.id === exception.signatureId)?.name, placement: exception.placement })),
      };
      outputZip.file("ASSINA-LOTE-MANIFESTO.json", JSON.stringify(manifest, null, 2));
      outputZip.file("LEIA-ME.txt", `Assina Lote\n\nDOCX: busca/substituição aplicada em document.xml, cabeçalhos e rodapés; assinatura visual posicionada nas páginas selecionadas. PDF: exportação e assinatura PNG aplicadas nas páginas; substituição textual em PDF requer OCR para preservar o layout. Busca: ${searchText || "(não aplicada)"}.\n`);
      const zipBytes = await outputZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
      downloadBytes(zipBytes, "documentos-assinados-assina-lote.zip", "application/zip");
      setLastExport(exported);
      setMessage(`${exported.length} documentos exportados. O download do ZIP começou.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Não foi possível exportar o lote.");
    } finally {
      setIsExporting(false);
    }
  };

  const activePageException = activePreviewDoc && activeSignature ? getException(activePreviewDoc.id, activePreviewPage, activeSignature.id) : undefined;
  const activePlacement = activePageException?.placement ?? activeSignature?.placement ?? DEFAULT_PLACEMENT;

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <div className="brand-mark"><PenLine size={21} strokeWidth={2.5} /></div>
          <div><div className="brand-name">Assina Lote</div><div className="brand-tag">documentos em movimento</div></div>
        </div>
        <div className="topbar-meta"><span className="secure-dot" /> processamento local no navegador <span className="topbar-sep" /> <LockKeyhole size={15} /> seus arquivos não saem desta aba</div>
      </header>

      <main className="workspace">
        <aside className="sidebar">
          <div className="eyebrow">FLUXO DE TRABALHO</div>
          <div className="step-list">
            {[
              { number: "01", label: "Enviar ZIP", Icon: FileArchive, done: documents.length > 0 },
              { number: "02", label: "Preparar assinaturas", Icon: PenLine, done: signatures.length > 0 },
              { number: "03", label: "Definir placement", Icon: MousePointer2, done: Boolean(activeSignature) },
              { number: "04", label: "Aplicar e baixar", Icon: ArrowDownToLine, done: lastExport.length > 0 },
            ].map(({ number, label, Icon: StepIcon, done }) => (
              <div className={`step ${done ? "is-done" : ""}`} key={number}>
                <span className="step-number">{done ? <Check size={14} /> : number}</span>
                <StepIcon size={17} />
                <span>{label}</span>
              </div>
            ))}
          </div>
          <div className="sidebar-note"><Sparkles size={16} /><span>Feito para lotes grandes: uma única prévia, muitos documentos.</span></div>
          <div className="sidebar-footer"><ShieldCheck size={15} /><span>Manifesto SHA-256 incluído no ZIP final</span></div>
        </aside>

        <section className="content">
          <div className="hero-row"><div><div className="eyebrow amber">ASSINATURA EM LOTE</div><h1>Assine dezenas de documentos<br /><em>sem repetir o gesto.</em></h1><p className="hero-copy">Suba seu pacote, trate sua assinatura, defina uma vez onde ela deve aparecer e exporte tudo pronto.</p></div><div className="hero-stamp"><Fingerprint size={28} /><strong>100%</strong><span>local</span></div></div>

          <div className="status-line"><span className={`status-pip ${message.includes("Falha") || message.includes("não") ? "warning" : ""}`} />{message}</div>

          <section className="panel upload-panel">
            <div className="panel-heading"><div className="panel-index">01</div><div><h2>Envie o pacote de documentos</h2><p>Aceitamos ZIP com arquivos <strong>.DOCX</strong> e <strong>.PDF</strong>. A leitura e o processamento acontecem no navegador.</p></div></div>
            <input ref={zipInputRef} className="hidden-input" type="file" accept=".zip,application/zip" onChange={onZipChange} />
            <button className="dropzone" onClick={() => zipInputRef.current?.click()}><div className="drop-icon"><Upload size={23} /></div><div><strong>{documents.length ? "Trocar pacote ZIP" : "Clique para enviar o ZIP"}</strong><span>{documents.length ? `${documents.length} arquivos selecionados` : "ou arraste seu arquivo para esta área"}</span></div><ChevronRight size={19} className="drop-arrow" /></button>
            {documents.length > 0 && <div className="doc-summary"><div className="summary-number">{documents.length}</div><div><strong>documentos prontos para assinatura</strong><span>Formato reconhecido: Word Open XML · páginas lógicas detectadas</span></div><button className="ghost-button" onClick={() => setDocuments([])}><RotateCcw size={15} /> limpar</button></div>}
          </section>

          <section className="panel replace-panel">
            <div className="panel-heading compact"><div className="panel-index">02</div><div><h2>Buscar e substituir em lote</h2><p>Aplicado no XML de DOCX, incluindo cabeçalhos e rodapés. PDFs são preservados e assinados sem reescrever texto.</p></div></div>
            <div className="replace-grid"><label>Localizar<input value={searchText} onChange={(event) => setSearchText(event.target.value)} placeholder="Texto atual" /></label><label>Substituir por<input value={replaceText} onChange={(event) => setReplaceText(event.target.value)} placeholder="Novo texto" /></label><button className="small-primary" onClick={() => void applyTextReplacement()} disabled={isApplyingReplacement}>{isApplyingReplacement ? "aplicando..." : "aplicar substituição"}<Check size={14} /></button></div>
          </section>

          <div className="two-col">
            <section className="panel signatures-panel">
              <div className="panel-heading compact"><div className="panel-index">02</div><div><h2>Prepare suas assinaturas</h2><p>Use uma imagem ou desenhe diretamente na tela.</p></div></div>
              <input ref={imageInputRef} className="hidden-input" type="file" accept="image/png,image/jpeg,image/jpg" onChange={onImageChange} />
              <div className="signature-actions"><button className="action-card" onClick={() => imageInputRef.current?.click()}><div className="action-icon coral"><ImagePlus size={20} /></div><strong>Subir PNG / JPG</strong><span>remoção automática de fundo</span></button><button className="action-card" onClick={() => { const canvas = canvasRef.current; canvas?.scrollIntoView({ behavior: "smooth", block: "center" }); }}><div className="action-icon teal"><PenLine size={20} /></div><strong>Desenhar assinatura</strong><span>touch, mouse ou caneta</span></button></div>
              <div className="draw-box"><div className="draw-toolbar"><span><span className="tiny-dot" style={{ backgroundColor: drawColor }} /> área de desenho</span><div className="draw-tools"><input aria-label="Cor da assinatura" type="color" value={drawColor} onChange={(event) => setDrawColor(event.target.value)} /><button onClick={clearDraw} title="Limpar desenho"><Trash2 size={14} /></button></div></div><canvas ref={canvasRef} width={900} height={320} onPointerDown={startDraw} onPointerMove={moveDraw} onPointerUp={finishDraw} onPointerCancel={finishDraw} /><div className="draw-footer"><span>Faça um traço livre no quadro acima</span><button className="small-primary" onClick={saveDrawnSignature}><Plus size={14} /> salvar desenho</button></div></div>
              {signatures.length > 0 && <div className="signature-list">{signatures.map((signature) => <div className={`signature-row ${activeSignature?.id === signature.id ? "active" : ""}`} key={signature.id}><input type="checkbox" checked={selectedSignatureIds.includes(signature.id)} onChange={() => toggleSignature(signature.id)} /><button className="signature-select" onClick={() => setActiveSignatureId(signature.id)}><span className="signature-thumb"><img src={signature.dataUrl} alt="" /></span><span><strong>{signature.name}</strong><small>{signature.kind === "image" ? "imagem tratada" : "desenho digital"}</small></span></button><button className="icon-button danger" onClick={() => removeSignature(signature.id)}><Trash2 size={15} /></button></div>)}</div>}
            </section>

            <section className="panel placement-panel">
              <div className="panel-heading compact"><div className="panel-index">03</div><div><h2>Defina o placement</h2><p>Arraste a assinatura na página-modelo e replique.</p></div></div>
              <div className="preview-toolbar">
                <div className="preview-doc-picker">
                  {documents.map((doc) => <button key={doc.id} className={activePreviewDoc?.id === doc.id ? "preview-doc-tab active" : "preview-doc-tab"} onClick={() => { setActivePreviewDocId(doc.id); setActivePreviewPage(1); setPlacementMode("batch"); }}><FileText size={13} /> <span>{doc.fileName}</span></button>)}
                </div>
                <span className="rendered-count">{activePreviewDoc ? activePreviewDoc.kind === "pdf" ? "PDF · exportação original" : `${renderedPageCount || activePreviewDoc.pages} páginas reais` : "aguardando DOCX/PDF"}</span>
              </div>
              {activePreviewDoc ? <>
                <div className="preview-modebar">
                  <button className="mode-button" onClick={() => void applyTextReplacement()} disabled={isApplyingReplacement || !searchText}><Check size={14} /> {isApplyingReplacement ? "aplicando…" : "aplicar busca/substituição"}</button>
                  <button className={placementMode === "batch" ? "mode-button active" : "mode-button"} onClick={() => setPlacementMode("batch")}><Layers3 size={14} /> placement em lote</button>
                  <button className={placementMode === "exception" ? "mode-button exception active" : "mode-button exception"} onClick={createPageException} disabled={!activeSignature}><Plus size={14} /> {activePageException ? "editar exceção" : "criar exceção nesta página"}</button>
                  <button className="remove-exception" onClick={removePageException} disabled={!activePageException}><Trash2 size={13} /> remover</button>
                </div>
                <div ref={previewStylesRef} className="docx-style-container" />
                <div className="real-preview-stage">
                  <div ref={previewRef} className="docx-render-container" />
                  {renderedPages.map(({ page, top, left, width, height }) => <div key={page} className={page === activePreviewPage ? "real-page-overlay active" : "real-page-overlay"} style={{ top, left, width, height }} onClick={() => { setActivePreviewPage(page); setPlacementMode(getException(activePreviewDoc.id, page, activeSignature?.id) ? "exception" : "batch"); }}>
                    <button className="real-page-chip" onClick={(event) => { event.stopPropagation(); setActivePreviewPage(page); }}>{page === activePreviewPage ? "página" : "p."} {page}</button>
                    {selectedSignatureIds.map((signatureId) => {
                      const signature = signatures.find((item) => item.id === signatureId);
                      if (!signature) return null;
                      const placement = getEffectivePlacement(activePreviewDoc.id, page, signature);
                      return <div key={`${page}-${signature.id}`} className={`real-signature-overlay ${signature.id === activeSignature?.id ? "selected" : ""} ${isDraggingPreview && signature.id === activeSignature?.id ? "dragging" : ""}`} onPointerDown={(event) => { setActiveSignatureId(signature.id); startPreviewDrag(event); }} onPointerMove={signature.id === activeSignature?.id ? movePreviewDrag : undefined} onPointerUp={() => setIsDraggingPreview(false)} onPointerCancel={() => setIsDraggingPreview(false)} style={{ left: `${placement.x}%`, top: `${placement.y}%`, width: `${placement.width}%`, opacity: placement.opacity / 100 }}><img src={signature.dataUrl} alt={`Assinatura ${signature.name}`} /></div>;
                    })}
                    {exceptions.some((exception) => exception.docId === activePreviewDoc.id && exception.page === page) && <span className="exception-badge"><WandSparkles size={11} /> exceção</span>}
                  </div>)}
                  {!renderedPages.length && <div className="preview-loading"><FileCheck2 size={22} /><span>{activePreviewDoc.kind === "pdf" ? "PDF pronto para seleção e exportação." : "Renderizando a página real do DOCX..."}</span></div>}
                </div>
                <div className="page-jump" aria-label="Navegar entre páginas">{Array.from({ length: renderedPageCount || activePreviewDoc.pages }, (_, index) => index + 1).map((page) => <button key={page} className={page === activePreviewPage ? "page-jump-button active" : "page-jump-button"} onClick={() => { setActivePreviewPage(page); setPlacementMode(getException(activePreviewDoc.id, page, activeSignature?.id) ? "exception" : "batch"); }}>p. {page}</button>)}</div>
                <div className="preview-caption"><MousePointer2 size={14} /> DOCX real · clique numa página para criar uma exceção específica</div>
                {activeSignature ? <>
                  <div className="placement-context"><strong>{placementMode === "exception" ? `Exceção · ${activePreviewDoc.fileName} · página ${activePreviewPage}` : "Placement em lote"}</strong><span>{placementMode === "exception" ? "Somente esta página será alterada; as demais continuam seguindo o lote." : "A posição será usada em todas as páginas selecionadas, salvo onde houver exceção."}</span></div>
                  <div className="sliders"><label>X <input type="range" min="0" max="85" value={activePlacement.x} onChange={(event) => updatePlacement("x", Number(event.target.value))} /><output>{Math.round(activePlacement.x)}%</output></label><label>Y <input type="range" min="0" max="90" value={activePlacement.y} onChange={(event) => updatePlacement("y", Number(event.target.value))} /><output>{Math.round(activePlacement.y)}%</output></label><label>Tamanho <input type="range" min="10" max="55" value={activePlacement.width} onChange={(event) => updatePlacement("width", Number(event.target.value))} /><output>{Math.round(activePlacement.width)}%</output></label><label>Opacidade <input type="range" min="30" max="100" value={activePlacement.opacity} onChange={(event) => updatePlacement("opacity", Number(event.target.value))} /><output>{Math.round(activePlacement.opacity)}%</output></label></div>
                  <button className="outline-button full" onClick={applyPlacementToSelected} disabled={placementMode === "exception"}><Layers3 size={15} /> aplicar este placement às assinaturas selecionadas</button>
                </> : <div className="empty-mini"><PenLine size={17} /> Selecione ou adicione uma assinatura para posicionar.</div>}
              </> : <div className="empty-mini"><FileText size={19} /> Envie um ZIP para visualizar o documento real.</div>}
            </section>
          </div>

          <section className="panel pages-panel">
            <div className="panel-heading compact"><div className="panel-index">04</div><div><h2>Escolha documentos e páginas</h2><p>O placement será aplicado apenas nos documentos e páginas marcados.</p></div><button className="outline-button select-all" onClick={toggleAllPages}>{allPagesSelected ? "desmarcar páginas" : "marcar todas as páginas"}</button></div>
            {documents.length > 0 ? <div className="document-table"><div className="table-head"><span>documento</span><span>páginas lógicas</span><span>status</span></div>{documents.map((doc) => <div className={`document-row ${doc.selected ? "selected" : ""}`} key={doc.id}><label className="doc-name"><input type="checkbox" checked={doc.selected} onChange={() => toggleDocument(doc.id)} /><FileText size={17} /><span>{doc.fileName}</span></label><div className="page-pills">{Array.from({ length: doc.pages }, (_, i) => i + 1).map((page) => <button className={doc.selectedPages.includes(page) ? "page-pill active" : "page-pill"} key={page} onClick={() => togglePage(doc.id, page)}>p. {page}</button>)}</div><span className="row-status">{doc.selected && doc.selectedPages.length > 0 ? <><Check size={14} /> pronto</> : "ignorado"}</span></div>)}</div> : <div className="empty-docs"><FileArchive size={25} /><strong>Seu pacote aparecerá aqui</strong><span>Depois do upload, você poderá escolher páginas individualmente.</span></div>}
          </section>

          <section className="export-card"><div className="export-copy"><div className="export-icon"><Archive size={22} /></div><div><div className="eyebrow amber">PRONTO PARA O LOTE</div><h2>Aplicar assinaturas e baixar ZIP</h2><p>O pacote final inclui os DOCX assinados, um manifesto SHA-256 e um arquivo de instruções.</p></div></div><button className="export-button" disabled={isExporting} onClick={exportZip}>{isExporting ? "processando..." : "gerar ZIP assinado"}<ArrowDownToLine size={18} /></button></section>

          <section className="verification-note"><div className="verification-icon"><ShieldCheck size={19} /></div><div><strong>Sobre validade e verificação</strong><p>A imagem tratada funciona como assinatura eletrônica visual. O ZIP exportado recebe hashes SHA-256 para verificação de integridade. Para assinatura avançada ou qualificada com validade jurídica de mercado, é necessário usar certificado digital e uma autoridade certificadora compatível.</p></div><span className="not-qualified">não ICP-Brasil</span></section>
        </section>
      </main>
    </div>
  );
}
