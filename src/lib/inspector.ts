import JSZip from 'jszip';

export interface ShapeInfo {
  id: string;
  name: string;
  type: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}

export interface ContainerInfo {
  index: number;
  label: string;
  shapes: ShapeInfo[];
}

export interface InspectionResult {
  fileName: string;
  fileType: 'pptx' | 'docx';
  containers: ContainerInfo[];
  thumbnailUrl?: string;
  pageSize?: {
    width: number;
    height: number;
  };
}

export async function inspectFile(file: File): Promise<InspectionResult> {
  const data = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(data);
  const fileType = file.name.toLowerCase().endsWith('.pptx') ? 'pptx' : 'docx';

  // Try to find a thumbnail (standard location is docProps/thumbnail.jpeg or .png)
  let thumbnailUrl: string | undefined;
  
  // Potential thumbnail locations in OOXML files
  const thumbnailPaths = [
    'docProps/thumbnail.jpeg',
    'docProps/thumbnail.png',
    'docProps/thumbnail.jpg',
    'docProps/thumbnail.bmp'
  ];

  for (const path of thumbnailPaths) {
    const thumbnailFile = zip.file(path);
    if (thumbnailFile) {
      try {
        const blob = await thumbnailFile.async('blob');
        thumbnailUrl = URL.createObjectURL(blob);
        break;
      } catch (e) {
        console.warn(`Failed to load thumbnail from ${path}`, e);
      }
    }
  }

  // If no standard thumbnail, we can try to find images in the slide folder if it's a PPTX
  // but that's not really a thumbnail of the whole slide.

  let result: InspectionResult;
  if (fileType === 'pptx') {
    result = await inspectPptx(zip, file.name);
  } else {
    result = await inspectDocx(zip, file.name);
  }

  return { ...result, thumbnailUrl };
}

async function inspectPptx(zip: JSZip, fileName: string): Promise<InspectionResult> {
  const containers: ContainerInfo[] = [];
  
  // Find slide size in ppt/presentation.xml
  let pageSize: { width: number; height: number } | undefined;
  const presentationContent = await zip.file('ppt/presentation.xml')?.async('text');
  if (presentationContent) {
    const parser = new DOMParser();
    const presDoc = parser.parseFromString(presentationContent, 'text/xml');
    const sldSz = presDoc.querySelector('sldSz');
    if (sldSz) {
      const cx = parseInt(sldSz.getAttribute('cx') || '9144000');
      const cy = parseInt(sldSz.getAttribute('cy') || '6858000');
      // Convert EMUs to points (1 EMU = 1/12700 point, 72 points per inch)
      // Actually often easier to just work in EMUs or normalize to 1000
      pageSize = { width: cx, height: cy };
    }
  }

  // Find all slide files
  const slideFiles = Object.keys(zip.files).filter(path => 
    path.startsWith('ppt/slides/slide') && path.endsWith('.xml')
  ).sort((a, b) => {
    const numA = parseInt(a.match(/\d+/)?.[0] || '0');
    const numB = parseInt(b.match(/\d+/)?.[0] || '0');
    return numA - numB;
  });

  const parser = new DOMParser();

  for (let i = 0; i < slideFiles.length; i++) {
    const path = slideFiles[i];
    const content = await zip.file(path)?.async('text');
    if (!content) continue;

    const xmlDoc = parser.parseFromString(content, 'text/xml');
    const shapes: ShapeInfo[] = [];

    // Most common: p:sp (shape), p:pic (picture), p:graphicFrame (charts/tables)
    // We want to find the root of the elements to find their translaton/offset
    const shapeContainers = xmlDoc.querySelectorAll('sp, pic, graphicFrame, grpSp, cxnSp');
    
    shapeContainers.forEach(container => {
      const nvPr = container.querySelector('cNvPr');
      if (!nvPr) return;

      const name = nvPr.getAttribute('name');
      const id = nvPr.getAttribute('id') || Math.random().toString(36).substr(2, 9);
      if (name) {
        let type = 'Element';
        const tag = container.tagName.split(':').pop();
        if (tag === 'sp') type = 'Shape';
        else if (tag === 'pic') type = 'Picture';
        else if (tag === 'graphicFrame') type = 'Graphic/Table';
        else if (tag === 'grpSp') type = 'Group';
        else if (tag === 'cxnSp') type = 'Connector';

        // Try to find off (offset) and ext (extent) in a:xfrm
        const xfrm = container.querySelector('xfrm');
        let x, y, width, height;
        if (xfrm) {
          const off = xfrm.querySelector('off');
          const ext = xfrm.querySelector('ext');
          if (off) {
            x = parseInt(off.getAttribute('x') || '0');
            y = parseInt(off.getAttribute('y') || '0');
          }
          if (ext) {
            width = parseInt(ext.getAttribute('cx') || '0');
            height = parseInt(ext.getAttribute('cy') || '0');
          }
        }
        
        shapes.push({ id, name, type, x, y, width, height });
      }
    });

    containers.push({
      index: i + 1,
      label: `Slide ${i + 1}`,
      shapes
    });
  }

  return { fileName, fileType: 'pptx', containers, pageSize };
}

async function inspectDocx(zip: JSZip, fileName: string): Promise<InspectionResult> {
  const content = await zip.file('word/document.xml')?.async('text');
  if (!content) {
    throw new Error('Not a valid DOCX file (missing word/document.xml)');
  }

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(content, 'text/xml');
  const shapes: ShapeInfo[] = [];

  // In Word, shapes/images often use 'wp:docPr' for names
  const docPrNodes = xmlDoc.querySelectorAll('docPr');
  docPrNodes.forEach(node => {
    const name = node.getAttribute('name');
    const id = node.getAttribute('id') || Math.random().toString(36).substr(2, 9);
    if (name) {
      shapes.push({ id, name, type: 'Drawing/Image' });
    }
  });

  // Older VML shapes
  const vShapeNodes = xmlDoc.querySelectorAll('shape');
  vShapeNodes.forEach(node => {
    const id = node.getAttribute('id') || node.getAttribute('o:spid');
    const alt = node.getAttribute('alt') || node.getAttribute('title');
    if (id) {
       // VML doesn't always have a friendly name like Selection Pane, but ID is used
       shapes.push({ id, name: alt || id, type: 'VML Shape' });
    }
  });

  // Extract page size from sectPr (usually at end of document.xml)
  let pageSize: { width: number; height: number } | undefined;
  const pgSz = xmlDoc.querySelector('pgSz');
  if (pgSz) {
    const w = parseInt(pgSz.getAttribute('w') || '11906'); // DXA (1/1440 inch)
    const h = parseInt(pgSz.getAttribute('h') || '16838');
    // Convert DXA to EMUs for consistency (1 DXA = 635 EMUs)
    pageSize = { width: w * 635, height: h * 635 };
  }

  return {
    fileName,
    fileType: 'docx',
    containers: [{
      index: 1,
      label: 'Document Body',
      shapes
    }],
    pageSize
  };
}
