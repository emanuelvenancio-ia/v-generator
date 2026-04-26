import JSZip from 'jszip';
import { InspectionResult } from './inspector';

interface Mapping {
  [columnName: string]: string; // columnName -> shapeName
}

export async function generateMergedFile(
  modelFile: File,
  excelData: { columns: string[], rows: any[] },
  mappings: Mapping,
  fileType: 'pptx' | 'docx'
): Promise<Blob> {
  const modelData = await modelFile.arrayBuffer();
  const zip = await JSZip.loadAsync(modelData);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  if (fileType === 'pptx') {
    return mergePptx(zip, excelData.rows, mappings, parser, serializer);
  } else {
    return mergeDocx(zip, excelData.rows, mappings, parser, serializer);
  }
}

export async function generateIndividualFiles(
  modelFile: File,
  excelData: { columns: string[], rows: any[] },
  mappings: Mapping,
  fileType: 'pptx' | 'docx'
): Promise<{ fileName: string, blob: Blob }[]> {
  const modelData = await modelFile.arrayBuffer();
  const results: { fileName: string, blob: Blob }[] = [];
  
  for (let i = 0; i < excelData.rows.length; i++) {
    const row = excelData.rows[i];
    const zip = await JSZip.loadAsync(modelData);
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    
    let blob: Blob;
    if (fileType === 'pptx') {
      blob = await mergePptx(zip, [row], mappings, parser, serializer);
    } else {
      blob = await mergeDocx(zip, [row], mappings, parser, serializer);
    }
    
    // Attempt to find a name for the file from mappings or common columns
    const firstMappingCol = Object.keys(mappings)[0];
    const itemLabel = row[firstMappingCol] ? String(row[firstMappingCol]).replace(/[^a-z0-9]/gi, '_') : `Item_${i + 1}`;
    
    results.push({
      fileName: `${itemLabel}.${fileType}`,
      blob
    });
  }
  
  return results;
}

async function mergePptx(
  zip: JSZip,
  rows: any[],
  mappings: Mapping,
  parser: DOMParser,
  serializer: XMLSerializer
): Promise<Blob> {
  // We'll take the first slide as template
  const slide1Path = 'ppt/slides/slide1.xml';
  const slide1Content = await zip.file(slide1Path)?.async('text');
  if (!slide1Content) throw new Error('Não foi possível encontrar o slide base (slide1.xml)');

  // To do a proper merge in PPTX, we need to create N slides
  // and update the presentation.xml + relationships.
  // For simplicity in this tool, lets try to replace text in the EXISTING slides if they match
  // OR create a new slide for each row.
  
  // NOTE: Creating slides requires updating ppt/presentation.xml and ppt/_rels/presentation.xml.rels
  // This is complex. A simpler "Mail Merge" for a web tool might be just producing ONE slide per row
  // but if the user wants it all in one file, we must do the complex way.
  
  const newZip = new JSZip();
  // Copy all files from original zip
  for (const [path, file] of Object.entries(zip.files)) {
    if (!path.startsWith('ppt/slides/slide') || path === 'ppt/slides/slide1.xml') {
        const content = await file.async('blob');
        newZip.file(path, content);
    }
  }

  const presentationXml = await zip.file('ppt/presentation.xml')?.async('text');
  const presentationRels = await zip.file('ppt/_rels/presentation.xml.rels')?.async('text');
  
  if (!presentationXml || !presentationRels) throw new Error('Estrutura de PPTX inválida');

  const presDoc = parser.parseFromString(presentationXml, 'text/xml');
  const relsDoc = parser.parseFromString(presentationRels, 'text/xml');
  
  const sldIdLst = presDoc.querySelector('sldIdLst');
  const relationships = relsDoc.querySelector('Relationships');
  
  if (!sldIdLst || !relationships) throw new Error('Estrutura de PPTX inconsistente');

  // Clear existing slide IDs and rels (except for slide 1 if we want to keep it as base, but we will replace it)
  // Actually let's just use slide1 as template and remove others
  while (sldIdLst.firstChild) sldIdLst.removeChild(sldIdLst.firstChild);
  
  // Keep only non-slide rels
  const relNodes = Array.from(relationships.querySelectorAll('Relationship'));
  relNodes.forEach(rel => {
    if (rel.getAttribute('Target')?.includes('slides/slide')) {
      relationships.removeChild(rel);
    }
  });

  // Now create a slide for each row
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const slideId = i + 256; // Standard starting ID
    const rId = `rIdMerge${i}`;
    const slidePath = `ppt/slides/slide${i + 1}.xml`;
    const slideRelPath = `ppt/slides/_rels/slide${i + 1}.xml.rels`;
    
    // Create Slide ID in presentation.xml
    const sldId = presDoc.createElementNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'p:sldId');
    sldId.setAttribute('id', slideId.toString());
    sldId.setAttribute('r:id', rId);
    sldIdLst.appendChild(sldId);
    
    // Create Relationship in presentation.xml.rels
    const rel = relsDoc.createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship');
    rel.setAttribute('Id', rId);
    rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
    rel.setAttribute('Target', `slides/slide${i + 1}.xml`);
    relationships.appendChild(rel);

    // Create the slide content
    let slideXml = parser.parseFromString(slide1Content, 'text/xml');
    
    // Replace text in shapes based on mappings
    Object.entries(mappings).forEach(([colName, shapeName]) => {
      const value = String(row[colName] || '');
      
      // Find shape with this name
      const shapes = Array.from(slideXml.querySelectorAll('sp, pic, graphicFrame'));
      shapes.forEach(shape => {
        const nvPr = shape.querySelector('cNvPr');
        if (nvPr?.getAttribute('name') === shapeName) {
          // --- CCUL 2026 Logic (inspired by VBA) ---
          let color: string | null = null;
          let labelText: string = value;
          
          if (colName === 'tipo_inscricao' || shapeName === 'txtTipo') {
            const raw = value.toUpperCase();
            if (raw.includes('MINI-CURSO')) color = '0066CC'; // Blue
            else if (raw.includes('AUTOR')) color = 'CC0000'; // Red
            else if (raw.includes('PRELECTOR')) color = '660099'; // Purple
            else if (raw.includes('EMPREENDEDOR')) color = 'FF8000'; // Orange
            else color = '00994C'; // Green
          }

          // Apply font color if it's a shape with text and we have a specific color
          if (color) {
            const txBody = shape.querySelector('txBody');
            if (txBody) {
              const rPrs = Array.from(txBody.querySelectorAll('rPr'));
              rPrs.forEach(rPr => {
                let solidFill = rPr.querySelector('solidFill');
                if (!solidFill) {
                  solidFill = slideXml.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:solidFill');
                  rPr.appendChild(solidFill);
                }
                while (solidFill.firstChild) solidFill.removeChild(solidFill.firstChild);
                const srgbClr = slideXml.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:srgbClr');
                srgbClr.setAttribute('val', color!);
                solidFill.appendChild(srgbClr);
              });
            }
          }

          // --- Standard Text Replacement ---
          const txBody = shape.querySelector('txBody');
          if (txBody) {
             const tNodes = Array.from(txBody.querySelectorAll('t'));
             if (tNodes.length > 0) {
               tNodes[0].textContent = labelText;
               for (let k = 1; k < tNodes.length; k++) {
                 tNodes[k].textContent = '';
               }
             }
          }
        }
      });
    });

    newZip.file(slidePath, serializer.serializeToString(slideXml));
    
    // Also need to copy slide1 rels if they exist
    const slide1RelPath = 'ppt/slides/_rels/slide1.xml.rels';
    const s1RelContent = await zip.file(slide1RelPath)?.async('blob');
    if (s1RelContent) {
      newZip.file(slideRelPath, s1RelContent);
    }
  }

  newZip.file('ppt/presentation.xml', serializer.serializeToString(presDoc));
  newZip.file('ppt/_rels/presentation.xml.rels', serializer.serializeToString(relsDoc));

  return await newZip.generateAsync({ type: 'blob' });
}

async function mergeDocx(
  zip: JSZip,
  rows: any[],
  mappings: Mapping,
  parser: DOMParser,
  serializer: XMLSerializer
): Promise<Blob> {
  // Word is a single file (document.xml)
  // For Word "Mail Merge", we usually repeat sections.
  // This is much harder to do genericly without knowing the structure.
  // One way is to just repeat the whole document body per row with page breaks.
  
  const docPath = 'word/document.xml';
  const docContent = await zip.file(docPath)?.async('text');
  if (!docContent) throw new Error('Não foi possível encontrar word/document.xml');

  const xmlDoc = parser.parseFromString(docContent, 'text/xml');
  const body = xmlDoc.querySelector('body');
  if (!body) throw new Error('Estrutura de DOCX inválida');

  // Keep a template of the body children except for the last sectPr
  const sectPr = body.querySelector('sectPr');
  const templateNodes = Array.from(body.childNodes).filter(node => node !== sectPr);
  
  while (body.firstChild) body.removeChild(body.firstChild);

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    
    // Clone nodes
    const itemNodes = templateNodes.map(node => node.cloneNode(true) as Element);
    
    // Replace text in shapes/drawings in these nodes
    itemNodes.forEach(node => {
      if (node.nodeType === 1) { // Element
        Object.entries(mappings).forEach(([colName, shapeName]) => {
          const value = String(row[colName] || '');
          
          // Find docPr with this name
          const docPrs = Array.from(node.querySelectorAll('docPr'));
          docPrs.forEach(docPr => {
            if (docPr.getAttribute('name') === shapeName) {
              // Now find the text to replace. In Word, drawings don't always have easy text.
              // BUT if we are mapping a column to a placeholder NAME, maybe the user wants 
              // to replace text inside the paragraph that has that drawing?
              // The request says "substituindo os shapes", which in Word is often textboxes.
              
              // Let's look for textboxes (w:txbxContent) inside the parent of this drawing
              let parent = docPr.parentElement;
              while (parent && parent.tagName !== 'w:p') parent = parent.parentElement;
              
              if (parent) {
                const tNodes = Array.from(parent.querySelectorAll('t'));
                if (tNodes.length > 0) {
                  tNodes[0].textContent = value;
                  for (let k = 1; k < tNodes.length; k++) tNodes[k].textContent = '';
                }
              }
            }
          });
          
          // Also look for direct text markers if we used VML IDs
          const vShapes = Array.from(node.querySelectorAll('shape'));
          vShapes.forEach(shape => {
            if (shape.getAttribute('id') === shapeName || shape.getAttribute('alt') === shapeName) {
                const tNodes = Array.from(shape.querySelectorAll('t'));
                if (tNodes.length > 0) {
                    tNodes[0].textContent = value;
                    for (let k = 1; k < tNodes.length; k++) tNodes[k].textContent = '';
                }
            }
          });
        });
      }
      body.appendChild(node);
    });

    // Add a page break between records (except last one maybe)
    if (i < rows.length - 1) {
      const p = xmlDoc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:p');
      const r = xmlDoc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
      const br = xmlDoc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:br');
      br.setAttribute('w:type', 'page');
      r.appendChild(br);
      p.appendChild(r);
      body.appendChild(p);
    }
  }

  // Restore sectPr
  if (sectPr) body.appendChild(sectPr);

  const newZip = new JSZip();
  for (const [path, file] of Object.entries(zip.files)) {
    if (path === docPath) {
      newZip.file(path, serializer.serializeToString(xmlDoc));
    } else {
      const content = await file.async('blob');
      newZip.file(path, content);
    }
  }

  return await newZip.generateAsync({ type: 'blob' });
}
