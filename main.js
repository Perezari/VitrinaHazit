window.jsPDF = window.jspdf.jsPDF;
function mm(v) { return Number.isFinite(v) ? v : 0; }

// ==== פרמטרים להתאמה מהירה ====
const PDF_ORIENTATION = 'l';   // landscape
const PDF_SIZE = 'a4';
const PAGE_MARGIN_MM = 1;     // מרווח למסך/תצוגה
const EXTRA_SCALE = 1.30;  // הגדלה לתצוגה

// ==== פרמטרים ייעודיים להדפסה (PDF) ====
const PRINT_MIN_MARGIN_MM = 10;   // שוליים בטוחים להדפסה
const PRINT_SAFE_SHRINK = 0.92; // כיווץ קל כדי למנוע חיתוך בקצה הנייר
const PRINT_ALIGN = 'center'; // 'left' | 'center'

// ==== גודל קבוע למסגרות ההערה (צהובות) ====
// שחק עם הערכים עד שזה עוטף את כל הטקסט בהדפסה:
const NOTE_BOX_W = 430;  // רוחב בפיקסלים-סגוליים של ה-SVG
const NOTE_BOX_H = 30;   // גובה בפיקסלים-סגוליים של ה-SVG

// ===== עזרי פונט עברית =====
function ensureAlefFont(pdf) {
    try {
        const list = pdf.getFontList ? pdf.getFontList() : null;
        const hasAlef = !!(list && (list.Alef || list['Alef']));
        if (hasAlef) { pdf.setFont('Alef', 'normal'); return; }
    } catch (_) { }
    if (typeof window.registerAlefFontOn === 'function') {
        const ok = window.registerAlefFontOn(pdf);
        if (ok) { pdf.setFont('Alef', 'normal'); return; }
    }
    if (typeof alefBase64 === 'string' && alefBase64.length > 100) {
        try {
            pdf.addFileToVFS('Alef-Regular.ttf', alefBase64);
            pdf.addFont('Alef-Regular.ttf', 'Alef', 'normal');
            pdf.setFont('Alef', 'normal');
            return;
        } catch (e) { console.warn('Font registration from base64 failed:', e); }
    }
    console.warn('Alef font not found; Hebrew may not render correctly.');
}

function withTempInDOM(svgNode, work) {
    const holder = document.createElement('div');
    holder.style.position = 'fixed';
    holder.style.left = '-10000px';
    holder.style.top = '-10000px';
    holder.style.opacity = '0';
    document.body.appendChild(holder);
    holder.appendChild(svgNode);
    try { return work(svgNode); }
    finally { document.body.removeChild(holder); }
}

function expandViewBoxToContent(svg, padding = 8) {
    const bbox = svg.getBBox();
    const minX = Math.floor(bbox.x - padding);
    const minY = Math.floor(bbox.y - padding);
    const width = Math.ceil(bbox.width + 2 * padding);
    const height = Math.ceil(bbox.height + 2 * padding);
    svg.setAttribute('viewBox', `${minX} ${minY} ${width} ${height}`);
}

// ===== המרת CSS לערכי attributes (stroke/dash/fill וכו׳) =====
function inlineComputedStyles(svgRoot) {
    svgRoot.querySelectorAll('*').forEach(el => {
        const cs = window.getComputedStyle(el);

        // טקסטים — פונט
        if (!el.getAttribute('font-family')) el.setAttribute('font-family', 'Alef');

        // קו/מילוי
        const stroke = cs.stroke && cs.stroke !== 'none' ? cs.stroke : null;
        const strokeWidth = cs.strokeWidth && cs.strokeWidth !== '0px' ? parseFloat(cs.strokeWidth) : null;
        const dash = cs.strokeDasharray && cs.strokeDasharray !== 'none' ? cs.strokeDasharray : null;
        const fillCss = cs.fill && cs.fill !== 'rgba(0, 0, 0, 1)' ? cs.fill : null;

        if (stroke) el.setAttribute('stroke', stroke);
        if (strokeWidth) el.setAttribute('stroke-width', strokeWidth);
        if (dash) el.setAttribute('stroke-dasharray', dash);

        // עובי קו קבוע למרות סקייל
        el.setAttribute('vector-effect', 'non-scaling-stroke');

        // קווי מידות
        if (el.classList && el.classList.contains('dim')) {
            if (!el.getAttribute('stroke')) el.setAttribute('stroke', '#2c3e50');
            if (!el.getAttribute('stroke-width')) el.setAttribute('stroke-width', '0.6');
        }

        // טיפול במילוי — לא לדרוס fill קיים שכבר הוגדר על האלמנט (למשל לעיגולים כחולים)
        if (['rect', 'path', 'polygon', 'polyline', 'circle', 'ellipse'].includes(el.tagName)) {
            if (el.classList && el.classList.contains('note-box')) {
                if (fillCss) el.setAttribute('fill', fillCss);
                if (!el.getAttribute('stroke')) el.setAttribute('stroke', '#2c3e50');
                if (!el.getAttribute('stroke-dasharray') && dash) el.setAttribute('stroke-dasharray', dash);
            } else {
                if (fillCss) {
                    el.setAttribute('fill', fillCss);
                } else if (!el.hasAttribute('fill')) {
                    el.setAttribute('fill', 'none');
                }
            }
        }
    });
}

// ===== תיקון טקסט עברית (bidi-override) =====
function fixHebrewText(svgRoot) {
    const hebrewRegex = /[\u0590-\u05FF]/;
    svgRoot.querySelectorAll('text').forEach(t => {
        const txt = (t.textContent || '').trim();
        if (!txt) return;
        if (hebrewRegex.test(txt)) {
            const reversed = txt.split('').reverse().join('');
            t.textContent = reversed;
            t.setAttribute('direction', 'ltr');
            t.setAttribute('unicode-bidi', 'bidi-override');
            t.setAttribute('font-family', 'Alef');
        }
    });
    svgRoot.setAttribute('direction', 'rtl');
}

/**
 * מרכז מספרי מידות על הקווים האנכיים או אופקיים,
 * כולל תמיכה בטקסטים עם סיבוב (rotate)
 * -- הערות: שמרנו את השינויים שלך והוספנו אופסט קטן
 */
function centerDimensionNumbers(svgRoot) {
    const numRegex = /^[\d\s\.\-+×xX*]+(?:mm|מ"מ|)$/;

    // גודל הפונט אחיד
    const fontSize = 12;

    svgRoot.querySelectorAll('text').forEach(t => {
        const raw = t.textContent || '';
        const txt = raw.replace(/\s+/g, '');
        if (!txt) return;

        if (numRegex.test(txt)) {
            // מרכז אופקי
            t.setAttribute('text-anchor', 'middle');

            // גודל פונט אחיד
            t.setAttribute('font-size', fontSize);

            // בדיקה אם יש סיבוב
            const transform = t.getAttribute('transform');
            if (transform && transform.includes('rotate')) {
                const match = /rotate\(([-\d.]+),\s*([-\d.]+),\s*([-\d.]+)\)/.exec(transform);
                if (match) {
                    const xRot = parseFloat(match[2]);
                    const yRot = parseFloat(match[3]);
                    const angle = parseFloat(match[1]);

                    if (Math.abs(angle) === 90) {
                        // טקסט אנכי: הזזה אופקית למרכז
                        const bbox = t.getBBox();
                        t.setAttribute('x', xRot);
                        t.setAttribute('y', yRot + 5);
                    } else {
                        // טקסט אופקי: השארת מיקום בסיסי
                        const bbox = t.getBBox();
                        t.setAttribute('x', xRot);
                        t.setAttribute('y', yRot);
                    }
                }
            } else {
                // טקסט אופקי רגיל: מרכז אנכי לפי bbox
                const bbox = t.getBBox();
                const y = parseFloat(t.getAttribute('y') || '0');
                t.setAttribute('y', y);
            }
        }
    });
}

// ===== חיצים במקום markers (תמיכה טובה יותר ב-PDF) =====
function replaceMarkersWithTriangles(svgRoot) {
    const lines = svgRoot.querySelectorAll('line, path, polyline');
    lines.forEach(el => {
        const hasMarker = el.getAttribute('marker-start') || el.getAttribute('marker-end');
        if (!hasMarker) return;

        // נתמוך בעיקר ב-line
        if (el.tagName !== 'line') {
            el.removeAttribute('marker-start');
            el.removeAttribute('marker-end');
            return;
        }

        const x1 = parseFloat(el.getAttribute('x1') || '0');
        const y1 = parseFloat(el.getAttribute('y1') || '0');
        const x2 = parseFloat(el.getAttribute('x2') || '0');
        const y2 = parseFloat(el.getAttribute('y2') || '0');
        const stroke = el.getAttribute('stroke') || '#000';
        const sw = parseFloat(el.getAttribute('stroke-width') || '1');

        const addTri = (x, y, angleRad) => {
            const size = Math.max(2.5 * sw, 3);
            const a = angleRad, s = size;
            const p1 = `${x},${y}`;
            const p2 = `${x - s * Math.cos(a - Math.PI / 8)},${y - s * Math.sin(a - Math.PI / 8)}`;
            const p3 = `${x - s * Math.cos(a + Math.PI / 8)},${y - s * Math.sin(a + Math.PI / 8)}`;
            const tri = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
            tri.setAttribute('points', `${p1} ${p2} ${p3}`);
            tri.setAttribute('fill', stroke);
            tri.setAttribute('stroke', 'none');
            el.parentNode.insertBefore(tri, el.nextSibling);
        };

        const ang = Math.atan2(y2 - y1, x2 - x1);
        if (el.getAttribute('marker-start')) addTri(x1, y1, ang + Math.PI);
        if (el.getAttribute('marker-end')) addTri(x2, y2, ang);

        el.removeAttribute('marker-start');
        el.removeAttribute('marker-end');
    });
}

// ===== חישוב התאמה + מיקום (יישור לשמאל/מרכז) =====
function fitAndPlaceBox(pdfWidth, pdfHeight, vbWidth, vbHeight, margin = 10, extraScale = 1.0, printShrink = 1.0, align = 'center') {
    const availW = pdfWidth - 2 * margin;
    const availH = pdfHeight - 2 * margin;
    const vbRatio = vbWidth / vbHeight;
    const pageRatio = availW / availH;

    let drawW, drawH;
    if (vbRatio > pageRatio) { drawW = availW; drawH = drawW / vbRatio; }
    else { drawH = availH; drawW = drawH * vbRatio; }

    drawW *= extraScale; drawH *= extraScale;
    drawW *= printShrink; drawH *= printShrink;

    // ביטחון נוסף
    if (drawW > pdfWidth - 2 * margin) { const s = (pdfWidth - 2 * margin) / drawW; drawW *= s; drawH *= s; }
    if (drawH > pdfHeight - 2 * margin) { const s = (pdfHeight - 2 * margin) / drawH; drawW *= s; drawH *= s; }

    // מיקום X לפי יישור
    let x;
    if (align === 'left') x = margin;                 // צמוד לשמאל (עד השוליים)
    else x = (pdfWidth - drawW) / 2; // מרכז

    const y = (pdfHeight - drawH) / 2; // נשאר ממורכז אנכית
    return { x, y, width: drawW, height: drawH };
}

/**
 * מכריח כל rect.note-box להיות בגודל קבוע NOTE_BOX_W × NOTE_BOX_H,
 * ממורכז סביב text.note-text שבאותו <g> — בלי לשנות stroke-width.
 */
function forceNoteBoxesSize(svgRoot, w = NOTE_BOX_W, h = NOTE_BOX_H) {
    const groups = svgRoot.querySelectorAll('g');
    groups.forEach(g => {
        const rect = g.querySelector('rect.note-box');
        const text = g.querySelector('text.note-text');
        if (!rect || !text) return;

        // --- טקסט ממורכז וקריא ---
        text.setAttribute('text-anchor', 'middle');
        text.setAttribute('dominant-baseline', 'middle');
        text.setAttribute('font-size', '17');
        if (!text.getAttribute('fill')) text.setAttribute('fill', '#111');

        // מבטיח שהטקסט נכנס יפה במסגרת
        const tb = text.getBBox();
        const cx = tb.x + tb.width / 2;
        const cy = tb.y + tb.height / 2;

        // קופסה קבועה סביב הטקסט
        const x = cx - w / 2 - 9;
        const y = cy - h / 2 - 5;

        rect.setAttribute('x', String(x));
        rect.setAttribute('y', String(y));
        rect.setAttribute('width', String(w));
        rect.setAttribute('height', String(h));

        // --- שיפורי נראות ---
        rect.setAttribute('rx', '6'); // פינות עגולות
        rect.setAttribute('ry', '6');
        rect.setAttribute('vector-effect', 'non-scaling-stroke');
        rect.setAttribute('shape-rendering', 'crispEdges');
        rect.setAttribute('fill-opacity', '0.9');

        // צבעים ברירת מחדל (אם לא קיימים)
        if (!rect.getAttribute('stroke')) rect.setAttribute('stroke', '#2c3e50');
        if (!rect.getAttribute('fill')) rect.setAttribute('fill', '#fff8b0');

        // דש-דש נשמר להדפסה
        const rc = getComputedStyle(rect);
        const dash = rc.strokeDasharray && rc.strokeDasharray !== 'none' ? rc.strokeDasharray : null;
        if (dash && !rect.getAttribute('stroke-dasharray')) {
            rect.setAttribute('stroke-dasharray', dash);
        }
    });

    // --- פילטר shadow (אם עדיין לא מוגדר) ---
    if (!svgRoot.querySelector('#noteBoxShadow')) {
        const defs = svgRoot.querySelector('defs') || svgRoot.insertBefore(document.createElementNS('http://www.w3.org/2000/svg', 'defs'), svgRoot.firstChild);
        const filter = document.createElementNS('http://www.w3.org/2000/svg', 'filter');
        filter.setAttribute('id', 'noteBoxShadow');
        filter.setAttribute('x', '-10%');
        filter.setAttribute('y', '-10%');
        filter.setAttribute('width', '120%');
        filter.setAttribute('height', '120%');
        const fe = document.createElementNS('http://www.w3.org/2000/svg', 'feDropShadow');
        fe.setAttribute('dx', '1');
        fe.setAttribute('dy', '1');
        fe.setAttribute('stdDeviation', '1');
        fe.setAttribute('flood-color', '#888');
        fe.setAttribute('flood-opacity', '0.5');
        filter.appendChild(fe);
        defs.appendChild(filter);
    }

    // להוסיף את הצל לכל note-box
    svgRoot.querySelectorAll('rect.note-box').forEach(r => {
        r.setAttribute('filter', 'url(#noteBoxShadow)');
    });
}

// נקודת מידה כחולה שלא נעלמת ב-PDF ולא משתנה בעובי
function addDimDot(svg, x, y, r = 2.2) {
    const c = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    c.setAttribute('cx', x);
    c.setAttribute('cy', y);
    c.setAttribute('r', r);
    // צבעי ברירת מחדל
    c.setAttribute('fill', '#54a5f5');
    c.setAttribute('stroke', '#2c3e50');
    // שומר על קו דק וחד בהדפסה
    c.setAttribute('stroke-width', '0.6');
    c.setAttribute('vector-effect', 'non-scaling-stroke');
    svg.appendChild(c);
}

// ====== פונקציית הייצוא ======
async function downloadPdf() {
    try {
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF(PDF_ORIENTATION, 'mm', PDF_SIZE);
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = pdf.internal.pageSize.getHeight();

        // קריאת נתוני היחידה
        const unitDetails = {
            sideSelect: document.getElementById('sideSelect').value,
            Sapak: document.getElementById('Sapak').value,
            planNum: document.getElementById('planNum').value,
            unitNum: document.getElementById('unitNum').value,
            partName: document.getElementById('partName').value,
            profileType: document.getElementById('profileType').value,
            profileColor: document.getElementById('profileColor').value,
            glassModel: document.getElementById('glassModel').value,
            glassTexture: document.getElementById('glassTexture').value,
            prepFor: document.getElementById('prepFor').value,
        };

        ensureAlefFont(pdf);

        // ====== טיפול ב-SVG ======
        const svgElement = document.getElementById('svg');
        if (!svgElement) { alert('לא נמצא אלמנט SVG לייצוא'); return; }
        const svgClone = svgElement.cloneNode(true);

        withTempInDOM(svgClone, (attached) => {
            inlineComputedStyles(attached);
            fixHebrewText(attached);
            centerDimensionNumbers(attached);
            replaceMarkersWithTriangles(attached);
            forceNoteBoxesSize(attached, NOTE_BOX_W, NOTE_BOX_H);
            expandViewBoxToContent(attached);
        });

        const vb2 = svgClone.viewBox && svgClone.viewBox.baseVal;
        const vbWidth = vb2 && vb2.width ? vb2.width : 1000;
        const vbHeight = vb2 && vb2.height ? vb2.height : 1000;

        const marginForPrint = Math.max(PAGE_MARGIN_MM, PRINT_MIN_MARGIN_MM);
        const displayExtra = Math.min(EXTRA_SCALE, 1.0);
        const box = fitAndPlaceBox(
            pdfWidth, pdfHeight, vbWidth, vbHeight,
            marginForPrint, displayExtra, PRINT_SAFE_SHRINK, PRINT_ALIGN
        );

        const options = { x: box.x, y: box.y, width: box.width, height: box.height, fontCallback: () => 'Alef' };
        let converted = false;

        if (typeof pdf.svg === 'function') {
            await pdf.svg(svgClone, options);
            converted = true;
        } else if (typeof window.svg2pdf === 'function') {
            await window.svg2pdf(svgClone, pdf, options);
            converted = true;
        }

        if (!converted) {
            const xml = new XMLSerializer().serializeToString(svgClone);
            const svg64 = window.btoa(unescape(encodeURIComponent(xml)));
            const imgSrc = 'data:image/svg+xml;base64,' + svg64;
            const img = new Image();
            img.crossOrigin = 'anonymous';
            await new Promise((res, rej) => { img.onload = res; img.onerror = rej; img.src = imgSrc; });
            pdf.addImage(img, 'PNG', box.x, box.y, box.width, box.height);
        }

        // ====== פרטי יחידה ======
        const textX = pdfWidth - marginForPrint;
        let textY = marginForPrint + 10;

        function fixHebrew(text) {
            return text.split('').reverse().join('');
        }

        function addFieldBox(label, value, width = 40, height = 10) {
            if (!value) return;
            pdf.setFont('Alef', 'normal');
            pdf.setFontSize(12);
            pdf.setTextColor(44, 62, 80);
            pdf.setFillColor(245);
            pdf.setDrawColor(200);
            pdf.setLineWidth(0.3);
            pdf.roundedRect(textX - width, textY, width, height, 3, 3, 'FD');

            const fixedValue = (label === 'מספר יחידה'
                || label === 'גוון פרופיל'
                || label === 'מספר תוכנית'
                || label === 'סוג זכוכית')
                ? value
                : fixHebrew(value);

            pdf.text(fixedValue, textX - width / 2, textY + height / 2, { align: 'center', baseline: 'middle' });

            const fixedLabel = fixHebrew(label);
            pdf.setFontSize(12);
            pdf.text(fixedLabel, textX - width / 2, textY - 1.5, { align: 'center' });

            textY += height + 7;
        }

        addFieldBox('הזמנה עבור', document.getElementById('Sapak').selectedOptions[0].text);
        addFieldBox('מספר תוכנית', unitDetails.planNum);
        addFieldBox('מספר יחידה', unitDetails.unitNum);
        addFieldBox('שם מפרק', unitDetails.partName);
        addFieldBox('סוג פרופיל', unitDetails.profileType);
        addFieldBox('גוון פרופיל', unitDetails.profileColor);
        addFieldBox('סוג זכוכית', unitDetails.glassModel);
        addFieldBox('כיוון טקסטורת זכוכית', document.getElementById('glassTexture').selectedOptions[0].text);
        addFieldBox('הכנה עבור', unitDetails.prepFor);

        // ====== הוספת לוגו לפי ספק ======
        function addLogo(pdf) {
            const supplier = unitDetails.Sapak;
            const logo = ProfileConfig.getLogoBySupplier("avivi");
            if (!supplier || !logo) return;

            const logoWidth = 40;
            const logoHeight = 25;
            pdf.addImage(logo, "PNG", 10, 10, logoWidth, logoHeight);
        }

        addLogo(pdf);

        function validateRequiredFields(fields) {
            let allValid = true;
            for (let id of fields) {
                const input = document.getElementById(id);
                if (input) {
                    if (input.value.trim() === '') {
                        alert('אנא מלא את השדה: ' + input.previousElementSibling.textContent);
                        input.style.border = '2px solid red';
                        input.focus();
                        allValid = false;
                        break; // עוצר בלחיצה הראשונה
                    } else {
                        // אם השדה לא ריק – מחזיר את העיצוב הרגיל
                        input.style.border = '';
                    }
                }
            }
            return allValid;
        }

        const requiredFields = [
            'Sapak',
            'planNum',
            'unitNum',
            'partName',
            'profileType',
            'profileColor',
            'glassModel',
        ];
        if (!validateRequiredFields(requiredFields)) return;

        // ====== שמירה ======
        function savePdf() {
            try {
                pdf.save(unitDetails.planNum + '_' + unitDetails.unitNum + '_' + unitDetails.profileType + '_' + unitDetails.sideSelect + '.pdf');
            } catch (_) {
                const blobUrl = pdf.output('bloburl');
                const a = document.createElement('a');
                a.href = blobUrl; a.download = 'שרטוט.pdf';
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                setTimeout(() => URL.revokeObjectURL(blobUrl), 1500);
            }
        }

        savePdf();

    } catch (err) {
        console.error('downloadPdf error:', err);
        alert('אירעה שגיאה בייצוא PDF. בדוק את הקונסול לפרטים.');
    }
}

function addNoteRotated(svg, x, y, text, angle = 90) {
    // מחשבים BBox זמני כדי להתאים את הריבוע
    const tempText = document.createElementNS("http://www.w3.org/2000/svg", "text");
    tempText.setAttribute("class", "note-text");
    tempText.setAttribute("x", x);
    tempText.setAttribute("y", y);
    tempText.setAttribute("text-anchor", "middle");
    tempText.setAttribute("dominant-baseline", "middle");
    tempText.textContent = text;
    svg.appendChild(tempText);

    const bbox = tempText.getBBox();
    svg.removeChild(tempText);

    const padding = 10;
    const rectX = bbox.x - padding;
    const rectY = bbox.y - padding;
    const rectW = bbox.width + padding * 2;
    const rectH = bbox.height + padding * 2;

    svg.insertAdjacentHTML("beforeend", `
    <g
    transform="rotate(${angle}, ${x}, ${y})">
    <rect class="note-box"
    x="${rectX}" 
    y="${rectY}"
    width="${rectW}" height="${rectH}"></rect>
    <text
    class="note-text"
    x="${x}" 
    y="${y}"
    text-anchor="middle"
    dominant-baseline="middle">
    ${text}
    </text>
    </g>
  `);
}

function draw() {
    const frontW = mm(+document.getElementById('frontW').value);
    const cabH = mm(+document.getElementById('cabH').value);
    const rEdge = mm(+document.getElementById('rEdge').value);
    const rMidCount = Math.max(1, mm(+document.getElementById('rMidCount').value) - 1);
    const rMidStep = (cabH - 2 * rEdge) / rMidCount;
    const scale = 0.16;
    const padX = 500, padY = 50;
    const W = frontW * scale, H = cabH * scale;
    const sideSelect = document.getElementById('sideSelect').value;
    const profileSelect = document.getElementById("profileType").value;


    const svg = document.getElementById('svg');
    const overlay = document.querySelector('.svg-overlay');
    overlay && (overlay.style.display = 'none');

    svg.innerHTML = `
    <defs>
    <marker
    id="arr"
    viewBox="0 0 10 10"
    refX="5"
    refY="5"
    markerWidth="0"
    markerHeight="0"
    orient="auto">
    <circle
    cx="5"
    cy="5"
    r="4"
    fill="#54a5f5"/>
    </marker>
    </defs>
    `;

    const profileType = document.getElementById('profileType').selectedOptions[0].text;
    const settings = ProfileConfig.getProfileSettings(profileType);

    const prepForInput = document.getElementById("prepFor");
    if (prepForInput) {
        prepForInput.value = settings.defaultPrepFor; // הערך האוטומטי
    }

    // חישוב מיקום וגודל פנימי
    const innerX = padX + settings.padSides * scale;
    const innerY = padY + settings.padTopBot * scale;
    const innerWW = W - 2 * settings.padSides * scale;
    const innerH = H - 2 * settings.padTopBot * scale;

    // מסגרת חיצונית
    svg.insertAdjacentHTML('beforeend',
        `<rect x="${padX}" 
         y="${padY}"
         width="${W}"
         height="${H}"
         fill="${settings.outerFrameFill}"
         stroke="${settings.outerFrameStroke}"
         stroke-width="${settings.outerFrameStrokeWidth}"/>`);

    //מסגרת פנימית
    svg.insertAdjacentHTML('beforeend',
        `<rect x="${innerX}"
         y="${innerY}"
         width="${innerWW}"
         height="${innerH}"
         fill="${settings.innerFrameFill}"
         stroke="${settings.innerFrameStroke}"
         stroke-width="${settings.innerFrameStrokeWidth}"/>`);

    if (settings.hasGerong) {
        // מסגרת עם גרונג - כמו בדגם קואדרו לדוגמה
        svg.insertAdjacentHTML('beforeend',
            `<line
             x1="${padX}" 
             y1="${padY}" 
             x2="${innerX}" 
             y2="${innerY}"
             stroke="#2c3e50"
             stroke-width="0.5"/>
             <line
             x1="${padX + W}" 
             y1="${padY}" 
             x2="${innerX + innerWW}" 
             y2="${innerY}" 
             stroke="#2c3e50" 
             stroke-width="0.5"/>
             <line
             x1="${padX}" 
             y1="${padY + H}" 
             x2="${innerX}" 
             y2="${innerY + innerH}"
             stroke="#2c3e50"
             stroke-width="0.5"/>
             <line
             x1="${padX + W}" 
             y1="${padY + H}" 
             x2="${innerX + innerWW}" 
             y2="${innerY + innerH}" 
             stroke="#2c3e50" 
             stroke-width="0.5"/>`
        );
    }
    else {
        //מסגרת ללא גרונג - כמו בדגם זירו לדוגמה
        svg.insertAdjacentHTML('beforeend',
            `<!-- קו עליון -->
            <line
            x1="${innerX - settings.padSides * scale}" 
            y1="${innerY}" 
            x2="${innerX + innerWW + settings.padSides * scale}" 
            y2="${innerY}" 
            stroke="#2c3e50" 
            stroke-width="0.5"/>
            <!-- קו תחתון -->
            <line
            x1="${innerX - settings.padSides * scale}" 
            y1="${innerY + innerH}" 
            x2="${innerX + innerWW + settings.padSides * scale}" 
            y2="${innerY + innerH}" 
            stroke="#2c3e50" 
            stroke-width="0.5"/>
            <!-- קו צד שמאל -->
            <line
            x1="${innerX}" 
            y1="${innerY}"
            x2="${innerX}"
            y2="${innerY + innerH}"
            stroke="#2c3e50"
            stroke-width="0.5"/>
            <!-- קו צד ימין -->
            <line
            x1="${innerX + innerWW}" 
            y1="${innerY}" 
            x2="${innerX + innerWW}" 
            y2="${innerY + innerH}" 
            stroke="#2c3e50" 
            stroke-width="0.5"/>`
        );
    }

    // קידוחים
    const drillR = 0.5; //קוטר קידוח
    const drillOffsetRight = 9.5; // הזזה אופקית לשרשרת ימין

    let frontDrillOffset = settings.frontDrillOffset;

    if (profileType === "ג'נסיס") {
        frontDrillOffset = frontDrillOffset * scale;
    }

    // --- קידוחים לאורך השרשרת הימנית (או שמאל לפי sideSelect) ---
    let yDrill = padY + rEdge * scale; // קידוח תחתון בדיוק בגובה rEdge
    let xRightDrill = padX + W - drillOffsetRight + frontDrillOffset;

    if (sideSelect === "left") {
        xRightDrill = padX + drillOffsetRight - frontDrillOffset;
    }

    let frontDimAdded = false; // דגל גלובלי למניעת חזרות

    function addDrill(x, y) {
        // קידוח רגיל
        svg.insertAdjacentHTML(
            'beforeend',
            `<circle cx="${x}" 
             cy="${y}" 
             r="${drillR}" 
             fill="none" 
             stroke="#2c3e50" 
             stroke-width="1"/>`
        );

        // אם זה פרופיל ג'נסיס – מוסיפים עוד אחד מעל (בהיסט של 32 מ״מ)
        if (settings.hasDualDrill) {
            const newY = y - settings.extraDrillOffset * scale;
            svg.insertAdjacentHTML(
                'beforeend',
                `<circle cx="${x}" 
                 cy="${newY}" 
                 r="${drillR}" 
                 fill="none" 
                 stroke="#2c3e50" 
                 stroke-width="1"/>`
            );

            if (!frontDimAdded) {
                // --- קו מקווקוו מקורי ---
                const lineYStart = newY;
                const lineYEnd = newY - 30;
                const lineX = (sideSelect === "left") ? padX : padX + W;
                const verticalLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
                verticalLine.setAttribute("x1", lineX);
                verticalLine.setAttribute("y1", lineYStart);
                verticalLine.setAttribute("x2", lineX);
                verticalLine.setAttribute("y2", lineYEnd);
                verticalLine.setAttribute("stroke", "#007acc");
                verticalLine.setAttribute("stroke-width", "1");
                verticalLine.setAttribute("stroke-dasharray", "4,2"); // קו מקווקוו
                svg.appendChild(verticalLine);

                // קו מקביל נוסף לקו המקווקוו
                let parallelX;
                if (sideSelect === "left") {
                    parallelX = lineX + frontDrillOffset / 2 - 1; // הזזה לכיוון הפנימי של הדלת
                } else {
                    parallelX = lineX - frontDrillOffset / 2 + 1; // הזזה לכיוון הפנימי של הדלת
                }

                const parallelLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
                parallelLine.setAttribute("x1", parallelX);
                parallelLine.setAttribute("y1", lineYStart);
                parallelLine.setAttribute("x2", parallelX);
                parallelLine.setAttribute("y2", lineYEnd);
                parallelLine.setAttribute("stroke", "#007acc");
                parallelLine.setAttribute("stroke-width", "1");
                parallelLine.setAttribute("stroke-dasharray", "4,2");
                svg.appendChild(parallelLine);

                // --- קו אופקי קצר לקידוח עם טקסט ---
                const horizontalLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
                horizontalLine.setAttribute("x1", lineX + 5);
                horizontalLine.setAttribute("y1", lineYEnd);
                horizontalLine.setAttribute("x2", x - 5);
                horizontalLine.setAttribute("y2", lineYEnd);
                horizontalLine.setAttribute("stroke", "#007acc");
                horizontalLine.setAttribute("stroke-width", "1");
                svg.appendChild(horizontalLine);

                // --- טקסט מידה מעל הקו ---
                const dimText = document.createElementNS("http://www.w3.org/2000/svg", "text");
                dimText.setAttribute("x", (lineX + x) / 2);
                dimText.setAttribute("y", lineYEnd - 5);
                dimText.setAttribute("text-anchor", "middle");
                dimText.setAttribute("class", "dim-text");
                dimText.textContent = settings.frontDrillOffset;
                svg.appendChild(dimText);

                const centerX = padX + W + 6; // מרכז רוחבי של החזית
                const lines = [settings.description];
                const fontSize = 10;
                const lineHeight = fontSize + 2;

                lines.forEach((line, i) => {
                    const textLine = document.createElementNS("http://www.w3.org/2000/svg", "text");
                    textLine.setAttribute("x", centerX);
                    textLine.setAttribute("y", padY + H + 20 + i * lineHeight); // מתחת לדלת
                    textLine.setAttribute("text-anchor", "middle");             // ממרכז את השורה אופקית
                    textLine.setAttribute("dominant-baseline", "middle");       // ממרכז את השורה אנכית
                    textLine.setAttribute("font-size", fontSize);
                    textLine.setAttribute("class", "dim-text");
                    textLine.textContent = line;
                    svg.appendChild(textLine);
                });
                frontDimAdded = true; // סמן שהקו נוסף
            }
        }
    }

    // קידוח ראשון
    addDrill(xRightDrill, yDrill);

    // קידוחים אמצעיים
    for (let i = 0; i < rMidCount; i++) {
        yDrill += rMidStep * scale;
        addDrill(xRightDrill, yDrill);
    }

    yDrill += rEdge * scale / 2;

    // ממדים ורוחב
    const dimY1 = padY + H + 30;
    svg.insertAdjacentHTML('beforeend', `<line 
                                         class="dim"
                                         x1="${padX}" 
                                         y1="${dimY1}" 
                                         x2="${padX + W}" 
                                         y2="${dimY1}">
                                         </line>`);
    // נקודות רוחב כולל
    addDimDot(svg, padX, dimY1);
    addDimDot(svg, padX + W, dimY1);
    svg.insertAdjacentHTML('beforeend', `<text 
                                         x="${padX + W / 2}" 
                                         y="${dimY1 + 16}" 
                                         text-anchor="middle">
                                         ${frontW}
                                         </text>`);

    // הגדרת מיקום הגובה הכולל לפי צד
    let xTotal;
    if (sideSelect === "right") {
        xTotal = padX - 30; // בצד ימין
    } else {
        xTotal = padX + W + 40;     // בצד שמאל
    }

    // ציור קו הגובה הכולל
    svg.insertAdjacentHTML('beforeend', `<line 
                                         class="dim" 
                                         x1="${xTotal - 10}" 
                                         y1="${padY}" 
                                         x2="${xTotal - 10}" 
                                         y2="${padY + H}">
                                         </line>`);
    // נקודות גובה כולל
    addDimDot(svg, xTotal - 10, padY);
    addDimDot(svg, xTotal - 10, padY + H);
    // תווית הגובה הכולל
    svg.insertAdjacentHTML('beforeend', `<text 
                                        x="${xTotal - 20}" 
                                        y="${padY + H / 2}" 
                                        transform="rotate(-90,${xTotal - 20},${padY + H / 2})" 
                                        text-anchor="middle" 
                                        dominant-baseline="middle">
                                        ${cabH}
                                        </text>`);

    // שרשראות ומדידות (ימין/שמאל)
    let xRightDim, xLeftDim;
    if (sideSelect === "right") {
        xRightDim = padX + W + 30;
        xLeftDim = padX - 30;
    } else {
        xRightDim = padX - 40;
        xLeftDim = padX + W + 40;
    }

    // שרשרת ימין
    let yR = padY;
    addDimDot(svg, xRightDim, yR);
    svg.insertAdjacentHTML('beforeend', `<line
                                         class="dim" 
                                         x1="${xRightDim}" 
                                         y1="${yR + 2}" 
                                         x2="${xRightDim}" 
                                         y2="${yR + rEdge * scale}">
                                         </line>`);

    addDimDot(svg, xRightDim, yR + (rEdge * scale));
    svg.insertAdjacentHTML('beforeend', `<text 
                                         x="${xRightDim + 20}" 
                                         y="${yR + (rEdge * scale) / 2 + 7}" 
                                         dominant-baseline="middle"
                                         transform="rotate(-90, ${xRightDim + 10}, ${yR + (rEdge * scale) / 2})">
                                         ${rEdge}
                                         </text>`);
    yR += rEdge * scale;

    for (let i = 0; i < rMidCount; i++) {
        // אם זו המידה שברצונך להסתיר
        if (i === 0) {
            yR += rMidStep * scale;        // רק עדכון y
            addDimDot(svg, xRightDim, yR); // נקודה נשמרת
            continue;
        }

        // קו
        svg.insertAdjacentHTML('beforeend', `<line 
                                             class="dim"
                                             x1="${xRightDim}"
                                             y1="${yR + 2}"
                                             x2="${xRightDim}"
                                             y2="${yR + rMidStep * scale}">
                                             </line>`);

        // נקודה
        addDimDot(svg, xRightDim, yR + (rMidStep * scale));

        // טקסט
        svg.insertAdjacentHTML('beforeend', `<text 
                                             x="${xRightDim + 20}"
                                             y="${yR + (rMidStep * scale) / 2 + 7}"
                                             dominant-baseline="middle"
                                             transform="rotate(-90, ${xRightDim + 10}, ${yR + (rMidStep * scale) / 2})">
                                             ${rMidStep.toFixed(0)}
                                             </text>`);

        yR += rMidStep * scale;
    }

    svg.insertAdjacentHTML('beforeend', `<line class="dim" x1="${xRightDim}" y1="${yR + 2}" x2="${xRightDim}" y2="${padY + H}"></line>`);
    addDimDot(svg, xRightDim, padY + H);
    svg.insertAdjacentHTML('beforeend', `<text x="${xRightDim + 20}" y="${yR + (padY + H - yR) / 2 + 7}" dominant-baseline="middle" transform="rotate(-90, ${xRightDim + 10}, ${yR + (padY + H - yR) / 2})">${rEdge}</text>`);

    const trueLeft = padX - 35;     // הקו הקיצוני בשמאל
    const trueRight = padX + W + 30; // הקו הקיצוני בימין
    const noteHeight = padY + H / 2;
    const noteOffset = 50;

    if (sideSelect === "right") {
        // מצב רגיל
        addNoteRotated(svg, trueRight + noteOffset, noteHeight, settings.rightNotes, 90);
    } else {
        // מצב שמאל – מתחלף
        addNoteRotated(svg, trueLeft - noteOffset, noteHeight, settings.rightNotes, -90);
    }

    // סיכום מידות
    const readoutContent = document.getElementById('readout-content');
    if (readoutContent) {
        readoutContent.innerHTML =
            `<div class="readout-item">גובה חזית: <strong>${cabH} מ״מ</strong></div>
             <div class="readout-item">רוחב חזית: <strong>${frontW} מ״מ</strong></div>
             <div class="readout-item">צירים: <strong>${rEdge} / ${rMidStep.toFixed(0)} / ${rEdge} מ״מ</strong></div>`;
        const readout = document.getElementById('readout');
        if (readout) readout.style.display = 'block';
    }
}

// חיבור כפתורים
const calcBtn = document.getElementById('calcBtn');
if (calcBtn) {
    calcBtn.addEventListener('click', () => {
        calcBtn.classList.add('loading');
        setTimeout(() => { draw(); calcBtn.classList.remove('loading'); }, 300);
    });
}

const downloadBtn = document.getElementById('downloadBtn');
if (downloadBtn) {
    downloadBtn.addEventListener('click', async (e) => {
        e.preventDefault();
        e.stopPropagation();
        try { await downloadPdf(); }
        catch (err) {
            console.error('[downloadPdf] failed:', err);
            alert('אירעה שגיאה בייצוא PDF. ראה קונסול.');
        }
    });
}

const sapakSelect = document.getElementById("Sapak");
const profileSelect = document.getElementById("profileType");
const sideSelect = document.getElementById("sideSelect");
const frontW = document.getElementById("frontW");
const cabH = document.getElementById("cabH");
const HingeLocation = document.getElementById("rEdge");
const HingeCount = document.getElementById("rMidCount");
const excelFile = document.getElementById("excelFile");
const unitContainer = document.getElementById("unitNum").parentElement;

let unitNumInput = document.getElementById("unitNum"); // משתנה שמצביע כרגע ל-input
let excelRows = []; // נשמור כאן את הנתונים מהקובץ

function fillProfileOptions() {
    const selectedSapak = sapakSelect.value;
    const options = ProfileConfig.getProfilesBySupplier(selectedSapak);

    profileSelect.innerHTML = "";

    options.forEach(profile => {
        const optionEl = document.createElement("option");
        optionEl.value = profile;
        optionEl.textContent = profile;
        profileSelect.appendChild(optionEl);
    });

    draw();
}

// מילוי ראשוני
fillProfileOptions();

// מאזינים לשינויים בזמן אמת
sapakSelect.addEventListener("change", fillProfileOptions);
profileSelect.addEventListener("change", draw);
sideSelect.addEventListener("change", draw);
frontW.addEventListener("change", draw);
cabH.addEventListener("change", draw);
HingeLocation.addEventListener("change", draw);
HingeCount.addEventListener("change", draw);

// טעינת קובץ Excel
excelFile.addEventListener("change", function (e) {
    const file = e.target.files[0];
    if (!file) return;

    // חילוץ מספר תוכנית מהשם
    const match = file.name.match(/^([A-Za-z0-9]+)_/);
    if (match) {
        document.getElementById('planNum').value = match[1];
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // range: 6 => להתחיל מהשורה 7 (B7) שבה יש כותרות
        excelRows = XLSX.utils.sheet_to_json(sheet, { range: 6 });

        console.log("עמודות שהתקבלו:", Object.keys(excelRows[0]));
        console.log("דוגמה לשורה ראשונה:", excelRows[0]);

        // הפיכת השדה unitNum לרשימה נפתחת אם הוא עדיין input
        if (unitNumInput.tagName.toLowerCase() === "input") {
            const select = document.createElement("select");
            select.id = "unitNum";

            // מספרי היחידות מהקובץ, מסוננים
            const units = [...new Set(
                excelRows
                    .map(r => String(r['יחידה']).trim())
                    .filter(u => u && u !== "undefined")
            )];

            units.forEach((unit, index) => {
                const option = document.createElement("option");
                option.value = unit;
                option.textContent = unit;
                select.appendChild(option);

                // בחר אוטומטית את הערך הראשון
                if (index === 0) select.value = unit;
            });

            // מחליפים את השדה ב-DOM
            unitContainer.replaceChild(select, unitNumInput);
            unitNumInput = select;
        }

        // מאזינים לשינוי ברשימה
        unitNumInput.addEventListener("change", function () {
            searchUnit(this.value);
        });

        // ניסיון ראשוני אם כבר יש מספר יחידה בשדה
        searchUnit(unitNumInput.value);
    };
    reader.readAsArrayBuffer(file);
});

// חיפוש שורה לפי מספר יחידה
function searchUnit(unitNum) {
    if (!excelRows.length || !unitNum) return;

    const row = excelRows.find(r => {
        const val = r['יחידה'];
        if (val === undefined) return false;
        return String(val).trim() === String(unitNum).trim();
    });

    if (!row) return;

    frontW.value = row['רוחב'] || '';
    cabH.value = row['אורך'] || '';

    // קביעת כיוון דלת לפי שם החלק
    if (row['שם החלק']) {
        const partName = row['שם החלק'].toLowerCase();
        if (partName.includes('ימין')) sideSelect.value = 'right';
        else if (partName.includes('שמאל')) sideSelect.value = 'left';
    }

    // סוג חומר -> גוון + סוג פרופיל
    if (row['סוג החומר']) {
        const [color, type] = row['סוג החומר'].split('_');
        document.getElementById('profileColor').value = color || '';

        // חיפוש ספק לפי סוג הפרופיל
        let foundSupplier = null;
        for (const supplier in ProfileConfig.SUPPLIERS_PROFILES_MAP) {
            if (ProfileConfig.SUPPLIERS_PROFILES_MAP[supplier].includes(type)) {
                foundSupplier = supplier;
                break;
            }
        }

        if (foundSupplier) {
            // עדכון הספק בשדה עם שם בעברית
            sapakSelect.value = foundSupplier;
            fillProfileOptions(); // עדכון הרשימה בהתאם לספק
        }

        profileSelect.value = type || '';
    }

    if (row['מלואה']) {
        document.getElementById('glassModel').value = row['מלואה'];
    }

    draw();
}

// חיפוש בלייב כשכותבים בשדה יחידה
unitNumInput.addEventListener("input", function () {
    searchUnit(this.value);
});

const batchSaveBtn = document.getElementById("batchSaveBtn");

function showOverlay() {
    const overlay = document.getElementById('overlay');
    overlay.style.display = 'flex';
    document.getElementById('overlayText').textContent = "שומר קבצים...";
    document.getElementById('overlayAnimation').textContent = "⏳";
}

function hideOverlayPending() {
    const overlay = document.getElementById('overlay');
    document.getElementById('overlayText').textContent = "קבצים נשלחו להורדה. אנא אשרו הורדות בדפדפן.";
    document.getElementById('overlayAnimation').textContent = "⬇️";
    setTimeout(() => {
        overlay.style.display = 'none';
    }, 3000); // 3 שניות לפני הסתרה
}

batchSaveBtn.addEventListener("click", async function () {
    if (!excelRows.length) return alert("אין קובץ Excel טעון!");

    showOverlay(); // מציג חלון המתנה

    // יצירת PDF לכל יחידה עם small delay כדי לייבא ערכים ל-DOM
    for (const row of excelRows) {
        if (!row['יחידה']) continue;

        const unitNumber = row['יחידה'];
        const partName = row['שם החלק'] || '';
        const material = row['סוג החומר'] || '';
        const glass = row['מלואה'] || '';

        // עדכון שדות כמו קודם
        frontW.value = row['רוחב'] || '';
        cabH.value = row['אורך'] || '';

        const doorSide = partName.includes('ימין') ? 'right' :
            partName.includes('שמאל') ? 'left' : '';
        sideSelect.value = doorSide;

        let profileType = '';
        let profileColor = '';
        if (material.includes('_')) [profileColor, profileType] = material.split('_');
        document.getElementById('profileColor').value = profileColor;

        let foundSupplier = null;
        for (const supplier in ProfileConfig.SUPPLIERS_PROFILES_MAP) {
            if (ProfileConfig.SUPPLIERS_PROFILES_MAP[supplier].includes(profileType)) {
                foundSupplier = supplier;
                break;
            }
        }
        if (foundSupplier) {
            sapakSelect.value = foundSupplier;
            fillProfileOptions();
        }

        profileSelect.value = profileType;
        document.getElementById('glassModel').value = glass;

        // עדכון שדה היחידה
        if (unitNumInput.tagName === 'SELECT') unitNumInput.value = unitNumber;
        else unitNumInput.value = unitNumber;

        const planNumber = document.getElementById('planNum').value;
        const fileName = `${planNumber}_${unitNumber}_${profileType}_${doorSide}.pdf`;

        // מחכה קצת בין קבצים כדי לעדכן DOM
        await new Promise(resolve => setTimeout(resolve, 50));

        generatePDFForUnit(fileName);
    }

    hideOverlayPending(); // מציג ✓ בסוף
});

function generatePDFForUnit(unitNumber) {
    // הפונקציה שלך שמייצרת PDF על פי הערכים הנוכחיים בשדות
    draw(); // אם צריך לעדכן את השרטוט לפני ההורדה
    // כאן הקוד ליצירת PDF והורדתו
    downloadPdf();
}

const excelFileInput = document.getElementById('excelFile');
const fileNameSpan = document.querySelector('.file-name');

excelFileInput.addEventListener('change', () => {
    if (excelFileInput.files.length > 0) {
        fileNameSpan.textContent = excelFileInput.files[0].name;
    } else {
        fileNameSpan.textContent = "לא נבחר קובץ";
    }
});

// הפעלה ראשונית
draw();