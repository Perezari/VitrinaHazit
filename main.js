window.jsPDF = window.jspdf.jsPDF;
function mm(v) { return Number.isFinite(v) ? v : 0; }

// Variables
const sapakSelect = document.getElementById("Sapak");
const profileSelect = document.getElementById("profileType");
const sideSelect = document.getElementById("sideSelect");
const frontW = document.getElementById("frontW");
const cabH = document.getElementById("cabH");
const HingeLocation = document.getElementById("rEdge");
const HingeCount = document.getElementById("rMidCount");
const excelFile = document.getElementById("excelFile");
const unitContainer = document.getElementById("unitNum").parentElement;
let unitNumInput = document.getElementById("unitNum");
let excelRows = [];
const downloadBtn = document.getElementById('downloadBtn');
const batchSaveBtn = document.getElementById("batchSaveBtn");
batchSaveBtn.style.display = 'none';
const excelFileInput = document.getElementById('excelFile');
const fileNameSpan = document.querySelector('.file-name');
let leftDoorWidth = 0;
let rightDoorWidth = 0;
let doorHeight = 0;

// משתנה גלובלי לשמירת המידות הניתנות לעריכה
let editableDimensions = {};

// CSS נוסף לעריכת מידות
const additionalCSS = `
.editable-dimension:hover {
    fill: #007acc !important;
    font-weight: bold;
}

.editable-dimension {
    transition: fill 0.2s ease;
    cursor: pointer;
}
`;

// הוספת ה-CSS לעמוד
const EditableDimensionstyle = document.createElement('style');
EditableDimensionstyle.textContent = additionalCSS;
document.head.appendChild(EditableDimensionstyle);

// פונקציה ליצירת מידה ניתנת לעריכה
function createEditableDimension(svg, x, y, value, id, rotation = 0, rotateX = null, rotateY = null, editable = true) {
    // שמירת הערך במשתנה הגלובלי
    editableDimensions[id] = value;

    const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
    text.setAttribute("x", x);
    text.setAttribute("y", y);
    text.setAttribute("text-anchor", "middle");
    text.setAttribute("dominant-baseline", "middle");

    if (rotation !== 0 && rotateX !== null && rotateY !== null) {
        text.setAttribute("transform", `rotate(${rotation}, ${rotateX}, ${rotateY})`);
    }

    text.setAttribute("class", "dim-text");
    text.setAttribute("data-dimension-id", id);
    text.setAttribute("style", "user-select: none;");

    text.textContent = value;

    // רק אם הפרמטר EDITABLE = true נוסיף את הקלאס ואת האירוע
    if (editable) {
        text.classList.add("editable-dimension");
        text.style.cursor = "pointer";

        // הוספת מאזין לדאבל קליק
        text.addEventListener("dblclick", function (e) {
            e.preventDefault();
            e.stopPropagation();
            editDimension(text, id);
        });
    }

    svg.appendChild(text);
    return text;
}

function editDimension(textElement, dimensionId) {
    const currentValue = editableDimensions[dimensionId];
    const rect = textElement.getBoundingClientRect();
    const svgRect = document.getElementById('svg').getBoundingClientRect();

    const input = document.createElement("input");
    input.type = "text";
    input.value = currentValue;
    input.style.position = "absolute";
    input.style.left = (rect.left - svgRect.left + 40) + "px";
    input.style.top = (rect.top - svgRect.top + 19) + "px";
    input.style.width = "45px";
    input.style.height = "25px";
    input.style.fontSize = "12px";
    input.style.textAlign = "center";
    input.style.border = "1px dashed #007acc";
    input.style.borderRadius = "4px";
    input.style.backgroundColor = "white";
    input.style.zIndex = "1000";
    input.style.direction = "ltr";

    const svgContainer = document.getElementById('svg').parentElement;
    svgContainer.appendChild(input);

    input.focus();
    input.select();

    function removeInput() {
        if (input && input.parentNode) {
            input.parentNode.removeChild(input);
        }
    }

    function saveValue() {
        const newValue = input.value.trim();
        if (newValue && !isNaN(newValue)) {
            editableDimensions[dimensionId] = newValue;
            textElement.textContent = newValue;
        }
        removeInput();
    }

    function cancelEdit() {
        removeInput();
    }

    input.addEventListener("keydown", function (e) {
        if (e.key === "Enter") {
            saveValue();
        } else if (e.key === "Escape") {
            cancelEdit();
        }
    });

    // כאן התיקון — דיליי קטן כדי שה־keydown יספיק לרוץ קודם
    input.addEventListener("blur", function () {
        setTimeout(() => {
            if (input && input.parentNode) {
                saveValue();
            }
        }, 0);
    });
}

// פונקציה לאיפוס המידות לערכים המקוריים (אופציונלי)
function resetDimensions() {
    editableDimensions = {};
    draw(); // צייר מחדש עם הערכים המקוריים
}

// Adds a small dot (circle) to the SVG at specified coordinates.
// Sets default fill and stroke colors for visibility.
// Ensures the stroke remains thin and sharp when scaling or printing.
// Default radius is 2.2 units, but can be overridden.
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

// Adds a rotated note box with text to the SVG.
// Temporarily measures the text to size the box with padding.
// Inserts a <g> element containing the <rect> and <text>, rotated around the specified coordinates.
// Default rotation angle is 90 degrees.
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
<g transform="rotate(${angle}, ${x}, ${y})"> <rect class="note-box" x="${rectX}" y="${rectY}" width="${rectW}" height="${rectH}"></rect> <text class="note-text" x="${x}" y="${y}" text-anchor="middle" dominant-baseline="middle"> ${text} </text> </g>
    `);
}

// Validates that all required input fields are filled.
// Shows an alert and highlights the first empty field.
// Returns true if all fields have values, false otherwise.
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

function validateRequiredFields(fields) {
    let allValid = true;
    let firstEmptyField = null;
    const inputs = [];

    // מעבר ראשון – לבדוק מי ריק
    for (let id of fields) {
        const input = document.getElementById(id);
        if (input) {
            inputs.push(input);
            if (input.value.trim() === '') {
                allValid = false;
                if (!firstEmptyField) {
                    firstEmptyField = input;
                }
            }
        }
    }

    // מעבר שני – צביעת שדות
    for (let input of inputs) {
        input.classList.remove('error', 'valid');
        if (input.value.trim() === '') {
            input.classList.add('error'); // שדות ריקים באדום
        }
    }

    if (firstEmptyField) {
        firstEmptyField.focus();
        showCustomAlert('אנא מלא את השדה: ' + firstEmptyField.previousElementSibling.textContent, "error");
    } else {
        showCustomAlert("מייצר PDF - אנא המתן", "success");
    }

    return allValid;
}

// פונקציה להצגת הודעה מותאמת אישית
function showCustomAlert(message, type = "error") {
    const alertDiv = document.createElement('div');
    alertDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 15px 25px;
        border-radius: 10px;
        font-weight: 600;
        z-index: 1000;
        animation: slideIn 0.3s ease;
    `;

    if (type === "error") {
        alertDiv.style.background = "linear-gradient(135deg, #ffcdd2 0%, #f8bbd9 100%)";
        alertDiv.style.color = "#c62828";
        alertDiv.style.boxShadow = "0 5px 15px rgba(198, 40, 40, 0.3)";
        alertDiv.style.borderRight = "4px solid #c62828";
    } else if (type === "success") {
        alertDiv.style.background = "linear-gradient(135deg, #c8e6c9 0%, #a5d6a7 100%)";
        alertDiv.style.color = "#2e7d32";
        alertDiv.style.boxShadow = "0 5px 15px rgba(46, 125, 50, 0.3)";
        alertDiv.style.borderRight = "4px solid #2e7d32";
    }

    // ספינר רק אם זה success
    if (type === "success") {
        const spinner = document.createElement("div");
        spinner.style.cssText = `
                width: 20px;
                height: 20px;
                border: 3px solid rgba(27, 94, 32, 0.3);
                border-top: 3px solid #1b5e20;
                border-radius: 50%;
                animation: spin 1s linear infinite;
                flex-shrink: 0;
        `;
        alertDiv.appendChild(spinner);
    }

    // מוחק את ההודעה אחרי 3 שניות (רק בשגיאה)
    if (type === "success") {
        setTimeout(() => {
            alertDiv.style.animation = 'slideOut 0.3s ease';
            setTimeout(() => {
                if (alertDiv.parentNode) {
                    alertDiv.parentNode.removeChild(alertDiv);
                }
            }, 300);
        }, 3000);
    }

    // טקסט
    const text = document.createElement("span");
    text.textContent = message;
    alertDiv.appendChild(text);

    document.body.appendChild(alertDiv);

    // מוחק את ההודעה אחרי 3 שניות (רק בשגיאה)
    if (type === "error") {
        setTimeout(() => {
            alertDiv.style.animation = 'slideOut 0.3s ease';
            setTimeout(() => {
                if (alertDiv.parentNode) {
                    alertDiv.parentNode.removeChild(alertDiv);
                }
            }, 300);
        }, 3000);
    }
}

// אנימציות CSS
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    @keyframes slideOut {
        from { transform: translateX(0); opacity: 1; }
        to { transform: translateX(100%); opacity: 0; }
    }
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
`;
document.head.appendChild(style);

// Populates the profile dropdown based on the selected supplier
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
fillProfileOptions();

// Finds a unit by number and fills the form fields with its properties
function searchUnit(unitNum) {
    if (!excelRows.length || !unitNum) return;

    // מסננים לפי מספר יחידה
    const candidates = excelRows.filter(r => {
        const val = r['יחידה'];
        return val !== undefined && String(val).trim() === String(unitNum).trim();
    });

    if (!candidates.length) return;

    // בדיקה אם יש גם דלת ימין וגם דלת שמאל
    const rightDoor = candidates.find(r => {
        const partName = (r['שם החלק'] || "").toLowerCase();
        return partName.includes("ימין");
    });

    const leftDoor = candidates.find(r => {
        const partName = (r['שם החלק'] || "").toLowerCase();
        return partName.includes("שמאל");
    });

    let row;

    // אם יש גם דלת ימין וגם דלת שמאל - מציבים "שתי דלתות"
    if (rightDoor && leftDoor) {
        sideSelect.value = 'both';

        // גובה יכול להיות זהה – או שאפשר לשמור גם שניים אם שונה
        cabH.value = rightDoor['אורך'] || leftDoor['אורך'] || '';

        // אם רוצים עדיין להמשיך לעבוד עם draw() כמו פעם,
        // אפשר לשמור גם ערך מחושב כללי, למשל:
        frontW.value = (Number(rightDoor['רוחב']) || 0) + (Number(leftDoor['רוחב']) || 0);

        // שמירה במשתנים
        leftDoorWidth = Number(leftDoor['רוחב']) || 0;
        rightDoorWidth = Number(rightDoor['רוחב']) || 0;
        doorHeight = Number(rightDoor['אורך']) || Number(leftDoor['אורך']) || 0;

        // ונבחר מה לשים בחומרים – למשל לפי דלת ימין
        row = rightDoor;
    } else if (rightDoor) {
        sideSelect.value = 'right';
        rightDoorWidth = Number(rightDoor['רוחב']) || 0;
        leftDoorWidth = 0;
        doorHeight = Number(rightDoor['אורך']) || 0;
        row = rightDoor;
    } else if (leftDoor) {
        sideSelect.value = 'left';
        leftDoorWidth = Number(leftDoor['רוחב']) || 0;
        rightDoorWidth = 0;
        doorHeight = Number(leftDoor['אורך']) || 0;
        row = leftDoor;
    } else {
        // לא נמצאה דלת - מציגים הודעה למשתמש
        alert("לא נמצאה דלת ימין או שמאל עבור יחידה " + unitNum);
        batchSaveBtn.disabled = true;
        return;
    }

    // אם נמצאה דלת – מחזירים את הכפתור לפעיל
    batchSaveBtn.disabled = false;

    // עדכון שדות
    frontW.value = row['רוחב'] || '';
    cabH.value = row['אורך'] || '';

    // סוג חומר -> גוון + סוג פרופיל
    if (row['סוג החומר']) {
        const [color, type] = row['סוג החומר'].split('_');
        document.getElementById('profileColor').value = color || '';

        let foundSupplier = null;
        for (const supplier in ProfileConfig.SUPPLIERS_PROFILES_MAP) {
            if (ProfileConfig.SUPPLIERS_PROFILES_MAP[supplier].includes(type)) {
                foundSupplier = supplier;
                break;
            }
        }

        if (foundSupplier) {
            sapakSelect.value = foundSupplier;
            fillProfileOptions();
        }

        profileSelect.value = type || '';
    }

    if (row['מלואה']) {
        document.getElementById('glassModel').value = row['מלואה'];
    }

    draw();
}

//Overlay functions
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
    setTimeout(() => { overlay.style.display = 'none'; }, 3000);
}

function generatePDFForUnit(unitNumber) {
    // הפונקציה שלך שמייצרת PDF על פי הערכים הנוכחיים בשדות
    draw(); // אם צריך לעדכן את השרטוט לפני ההורדה
    // כאן הקוד ליצירת PDF והורדתו
    downloadPdf();
}

// Generates a PDF from the current SVG and unit details on the page.
// Ensures Hebrew text uses the Alef font and applies all SVG styling and fixes.
// Clones the SVG, applies computed styles, fixes Hebrew text, centers dimensions, replaces markers, and sizes note boxes.
// Fits the SVG into the PDF page with proper scaling and margins.
// Adds unit detail fields as labeled boxes alongside the SVG.
// Adds supplier logos (PNG or SVG) to the PDF.
// Validates required fields before saving.
// Saves the PDF with a filename based on plan number, unit number, profile type, and side selection.
// Catches and reports errors during the PDF generation process.
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
            if (!text) return '';

            // זיהוי עברית
            const hebrewRegex = /[\u0590-\u05FF]/;

            if (hebrewRegex.test(text)) {
                // אם יש עברית – נהפוך את כל המחרוזת
                return text.split('').reverse().join('');
            }

            // אחרת אנגלית/מספרים – משאירים כמו שזה
            return text;
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

            const fixedValue = fixHebrew(value);
            const fixedLabel = fixHebrew(label);

            pdf.text(fixedValue, textX - width / 2, textY + height / 2, { align: 'center', baseline: 'middle' });

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
            const logo = ProfileConfig.getLogoBySupplier(supplier);
            if (!supplier || !logo) return;

            const logoWidth = 40;
            const logoHeight = 25;
            pdf.addImage(logo, "PNG", 10, 10, logoWidth, logoHeight);
        }

        async function addLogoSvg(pdf, logo) {
            if (!logo) return;

            let svgText;

            // בדיקה אם זה Data URI (base64)
            if (logo.startsWith("data:image/svg+xml")) {
                const base64 = logo.split(",")[1];
                svgText = atob(base64);
            } else {
                // SVG כטקסט רגיל
                svgText = logo;
            }

            // ממירים ל־DOM
            const svgElement = new DOMParser().parseFromString(svgText, "image/svg+xml").documentElement;

            // מוסיפים ל־PDF
            await pdf.svg(svgElement, {
                x: 10,
                y: 10,
                width: 40,
                height: 25
            });
        }

        // לוגו מ־ProfileConfig (יכול להיות טקסט או Data URI)
        const logo = ProfileConfig.getLogoBySupplier("avivi_svg");
        await addLogoSvg(pdf, logo);

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

// פונקציה לציור שרשרת מידות של קידוחים
function drawDrillChain(svg, padX, padY, W, H, rEdge, rMidCount, rMidStep, scale, side) {
    // קובע מיקום של השרשרת (ימין או שמאל)
    let xRightDim = side === "right" ? padX + W + 30 : padX - 40;

    let yR = padY;
    addDimDot(svg, xRightDim, yR);

    // מידה עליונה (rEdge)
    svg.insertAdjacentHTML('beforeend', `<line class="dim" 
        x1="${xRightDim}" y1="${yR + 2}" 
        x2="${xRightDim}" y2="${yR + rEdge * scale}"></line>`);
    addDimDot(svg, xRightDim, yR + (rEdge * scale));
    createEditableDimension(svg, xRightDim + 10, yR + (rEdge * scale) / 2 + 7, rEdge, `rEdge-top-${side}`, -90, xRightDim + 10, yR + (rEdge * scale) / 2);

    yR += rEdge * scale;

    // מידות ביניים (rMidStep)
    for (let i = 0; i < rMidCount; i++) {
        if (i === 0) {
            yR += rMidStep * scale;
            addDimDot(svg, xRightDim, yR);
            continue;
        }

        svg.insertAdjacentHTML('beforeend', `<line class="dim" 
            x1="${xRightDim}" y1="${yR + 2}" 
            x2="${xRightDim}" y2="${yR + rMidStep * scale}"></line>`);
        addDimDot(svg, xRightDim, yR + (rMidStep * scale));
        createEditableDimension(svg, xRightDim + 10, yR + (rMidStep * scale) / 2 + 7, rMidStep.toFixed(0), `rMid-${side}-${i}`, -90, xRightDim + 10, yR + (rMidStep * scale) / 2);

        yR += rMidStep * scale;
    }

    // מידה תחתונה (rEdge)
    svg.insertAdjacentHTML('beforeend', `<line class="dim" 
        x1="${xRightDim}" y1="${yR + 2}" 
        x2="${xRightDim}" y2="${padY + H}"></line>`);
    addDimDot(svg, xRightDim, padY + H);
    createEditableDimension(svg, xRightDim + 10, yR + (padY + H - yR) / 2 + 7, rEdge, `rEdge-bottom-${side}`, -90, xRightDim + 10, yR + (padY + H - yR) / 2);
}

// פונקציה לציור דלת יחידה
function drawSingleDoor(svg, padX, padY, W, H, side, settings, scale, frontW, cabH, rEdge, rMidCount, rMidStep, profileType, doorIndex = 0) {
    const drillR = 0.5;
    const drillOffsetRight = 9.5;

    let frontDrillOffset = settings.frontDrillOffset;
    if (profileType === "ג'נסיס") {
        frontDrillOffset = frontDrillOffset * scale;
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

    // מסגרת פנימית
    svg.insertAdjacentHTML('beforeend',
        `<rect x="${innerX}"
         y="${innerY}"
         width="${innerWW}"
         height="${innerH}"
         fill="${settings.innerFrameFill}"
         stroke="${settings.innerFrameStroke}"
         stroke-width="${settings.innerFrameStrokeWidth}"/>`);

    // קווי גרונג או רגילים
    if (settings.hasGerong) {
        svg.insertAdjacentHTML('beforeend',
            `<line x1="${padX}" y1="${padY}" x2="${innerX}" y2="${innerY}" stroke="#2c3e50" stroke-width="0.5"/>
             <line x1="${padX + W}" y1="${padY}" x2="${innerX + innerWW}" y2="${innerY}" stroke="#2c3e50" stroke-width="0.5"/>
             <line x1="${padX}" y1="${padY + H}" x2="${innerX}" y2="${innerY + innerH}" stroke="#2c3e50" stroke-width="0.5"/>
             <line x1="${padX + W}" y1="${padY + H}" x2="${innerX + innerWW}" y2="${innerY + innerH}" stroke="#2c3e50" stroke-width="0.5"/>`);
    } else {
        svg.insertAdjacentHTML('beforeend',
            `<line x1="${innerX - settings.padSides * scale}" y1="${innerY}" x2="${innerX + innerWW + settings.padSides * scale}" y2="${innerY}" stroke="#2c3e50" stroke-width="0.5"/>
             <line x1="${innerX - settings.padSides * scale}" y1="${innerY + innerH}" x2="${innerX + innerWW + settings.padSides * scale}" y2="${innerY + innerH}" stroke="#2c3e50" stroke-width="0.5"/>
             <line x1="${innerX}" y1="${innerY}" x2="${innerX}" y2="${innerY + innerH}" stroke="#2c3e50" stroke-width="0.5"/>
             <line x1="${innerX + innerWW}" y1="${innerY}" x2="${innerX + innerWW}" y2="${innerY + innerH}" stroke="#2c3e50" stroke-width="0.5"/>`);
    }

    // קידוחים
    let yDrill = padY + rEdge * scale;
    let xRightDrill = side === "left" ? padX + drillOffsetRight - frontDrillOffset : padX + W - drillOffsetRight + frontDrillOffset;

    let frontDimAdded = false;

    function addDrill(x, y) {
        let drillPositions = [];
        if (settings.hasDualDrill) {
            const halfOffset = (settings.extraDrillOffset / 2) * scale;
            drillPositions = [y - halfOffset, y + halfOffset];
        } else {
            drillPositions = [y];
        }

        drillPositions.forEach(posY => {
            svg.insertAdjacentHTML('beforeend',
                `<circle cx="${x}" cy="${posY}" r="${drillR}" fill="none" stroke="#2c3e50" stroke-width="1"/>`);
        });

        // הוספת קווים ומידות לקידוח כפול
        if (settings.hasDualDrill && !frontDimAdded) {
            const newY = y - settings.extraDrillOffset * scale;
            const lineYStart = newY;
            const lineYEnd = newY - 30;
            const lineX = side === "left" ? padX : padX + W;

            const verticalLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
            verticalLine.setAttribute("x1", lineX);
            verticalLine.setAttribute("y1", lineYStart);
            verticalLine.setAttribute("x2", lineX);
            verticalLine.setAttribute("y2", lineYEnd);
            verticalLine.setAttribute("stroke", "#007acc");
            verticalLine.setAttribute("stroke-width", "1");
            verticalLine.setAttribute("stroke-dasharray", "4,2");
            svg.appendChild(verticalLine);

            const parallelX = side === "left" ? lineX + frontDrillOffset / 2 - 1 : lineX - frontDrillOffset / 2 + 1;
            const parallelLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
            parallelLine.setAttribute("x1", parallelX);
            parallelLine.setAttribute("y1", lineYStart);
            parallelLine.setAttribute("x2", parallelX);
            parallelLine.setAttribute("y2", lineYEnd);
            parallelLine.setAttribute("stroke", "#007acc");
            parallelLine.setAttribute("stroke-width", "1");
            parallelLine.setAttribute("stroke-dasharray", "4,2");
            svg.appendChild(parallelLine);

            const horizontalLine = document.createElementNS("http://www.w3.org/2000/svg", "line");
            horizontalLine.setAttribute("x1", lineX + 5);
            horizontalLine.setAttribute("y1", lineYEnd);
            horizontalLine.setAttribute("x2", x - 5);
            horizontalLine.setAttribute("y2", lineYEnd);
            horizontalLine.setAttribute("stroke", "#007acc");
            horizontalLine.setAttribute("stroke-width", "1");
            svg.appendChild(horizontalLine);

            const dimText = document.createElementNS("http://www.w3.org/2000/svg", "text");
            dimText.setAttribute("x", (lineX + x) / 2);
            dimText.setAttribute("y", lineYEnd - 5);
            dimText.setAttribute("text-anchor", "middle");
            dimText.setAttribute("class", "dim-text");
            dimText.textContent = settings.frontDrillOffset;
            svg.appendChild(dimText);

            frontDimAdded = true;
        }
    }

    // קידוח ראשון
    addDrill(xRightDrill, yDrill);

    // קידוחים אמצעיים
    for (let i = 0; i < rMidCount; i++) {
        yDrill += rMidStep * scale;
        addDrill(xRightDrill, yDrill);
    }
}

// Draws a cabinet/front panel diagram in an SVG element.
// Includes frames, shelves, drill holes, dimensions, and rotated notes
// based on user input and profile settings.
// Also updates an HTML readout with the cabinet dimensions.
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

    if (sideSelect === 'both' && !excelRows.length) {  // אין קובץ אקסל
        const frontVal = mm(+document.getElementById('frontW').value);
        leftDoorWidth = frontVal;
        rightDoorWidth = frontVal;
    }

    // הגדרת viewBox לפי סוג הדלת
    let totalWidth, totalHeight;

    if (sideSelect === 'both') {
        // עבור שתי דלתות - רוחב כפול
        totalWidth = padX + (W * 2) + 100 + 480; // רווח בין הדלתות + מקום למידות
        totalHeight = padY + H + 150;
    } else {
        // עבור דלת יחידה
        totalWidth = padX + W + 480;
        totalHeight = padY + H + 150;
    }

    svg.setAttribute('viewBox', `0 0 ${totalWidth} ${totalHeight}`);
    svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');

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
        prepForInput.value = settings.defaultPrepFor;
    }

    if (sideSelect === 'both') {
        // ציור שתי דלתות
        const doorSpacing = 10; // רווח בין הדלתות
        const leftDoorX = padX;
        const rightDoorX = padX + (leftDoorWidth * 0.16) + doorSpacing;

        // ציור דלת שמאל
        drawSingleDoor(svg, leftDoorX, padY, leftDoorWidth * 0.16, H, 'left', settings, scale, leftDoorWidth, cabH, rEdge, rMidCount, rMidStep, profileType, 0);

        // ציור דלת ימין
        drawSingleDoor(svg, rightDoorX, padY, rightDoorWidth * 0.16, H, 'right', settings, scale, rightDoorWidth, cabH, rEdge, rMidCount, rMidStep, profileType, 1);

        // מידות עבור שתי דלתות
        const dimY1 = padY + H + 30;

        // רוחב דלת שמאל
        svg.insertAdjacentHTML('beforeend', `<line class="dim" x1="${leftDoorX}" y1="${dimY1}" x2="${leftDoorX + (leftDoorWidth * 0.16) }" y2="${dimY1}"></line>`);
        addDimDot(svg, leftDoorX, dimY1);
        addDimDot(svg, leftDoorX + (leftDoorWidth * 0.16), dimY1);
        createEditableDimension(svg, leftDoorX + (leftDoorWidth * 0.16) / 2, dimY1 + 16, leftDoorWidth, `frontW_left`, 0, null, null, false);

        // רוחב דלת ימין
        svg.insertAdjacentHTML('beforeend', `<line class="dim" x1="${rightDoorX}" y1="${dimY1}" x2="${rightDoorX + (rightDoorWidth * 0.16)}" y2="${dimY1}"></line>`);
        addDimDot(svg, rightDoorX, dimY1);
        addDimDot(svg, rightDoorX + (rightDoorWidth * 0.16), dimY1);
        createEditableDimension(svg, rightDoorX + (rightDoorWidth * 0.16) / 2, dimY1 + 16, rightDoorWidth, `frontW_right`, 0, null, null, false);

        // גובה כולל - צד שמאל
        const xTotal = leftDoorX - 30;
        svg.insertAdjacentHTML('beforeend', `<line class="dim" x1="${xTotal - 10}" y1="${padY}" x2="${xTotal - 10}" y2="${padY + H}"></line>`);
        addDimDot(svg, xTotal - 10, padY);
        addDimDot(svg, xTotal - 10, padY + H);
        createEditableDimension(svg, xTotal - 20, padY + H / 2, cabH, `cabH`, -90, xTotal - 20, padY + H / 2, null, null, false);

        drawDrillChain(svg, rightDoorX, padY, W, H, rEdge, rMidCount, rMidStep, scale, "right");

        // הערות
        const noteHeight = padY + H / 2;
        const noteOffset = 100;
        addNoteRotated(svg, leftDoorX - noteOffset, noteHeight, settings.rightNotes, -90);
    }

    else {
        // ציור דלת יחידה (הקוד המקורי)
        drawSingleDoor(svg, padX, padY, W, H, sideSelect, settings, scale, frontW, cabH, rEdge, rMidCount, rMidStep, profileType);

        // מידות רוחב
        const dimY1 = padY + H + 30;
        svg.insertAdjacentHTML('beforeend', `<line class="dim" x1="${padX}" y1="${dimY1}" x2="${padX + W}" y2="${dimY1}"></line>`);
        addDimDot(svg, padX, dimY1);
        addDimDot(svg, padX + W, dimY1);
        createEditableDimension(svg, padX + W / 2, dimY1 + 16, frontW, `frontW`, 0, null, null, false);

        // גובה כולל
        let xTotal = sideSelect === "right" ? padX - 30 : padX + W + 40;
        svg.insertAdjacentHTML('beforeend', `<line class="dim" x1="${xTotal - 10}" y1="${padY}" x2="${xTotal - 10}" y2="${padY + H}"></line>`);
        addDimDot(svg, xTotal - 10, padY);
        addDimDot(svg, xTotal - 10, padY + H);
        createEditableDimension(svg, xTotal - 20, padY + H / 2, cabH, `cabH`, -90, xTotal - 20, padY + H / 2, null, null, false);

        // שרשראות ומדידות
        drawDrillChain(svg, padX, padY, W, H, rEdge, rMidCount, rMidStep, scale, sideSelect);

        // הערות
        const trueLeft = padX - 40;
        const trueRight = padX + W + 30;
        const noteHeight = padY + H / 2;
        const noteOffset = 50;

        if (sideSelect === "right") {
            addNoteRotated(svg, trueRight + noteOffset, noteHeight, settings.rightNotes, 90);
        }
        else {
            addNoteRotated(svg, trueLeft - noteOffset, noteHeight, settings.rightNotes, -90);
        }
    }

    // סיכום מידות
    const readoutContent = document.getElementById('readout-content');
    if (readoutContent) {
        const doorType = sideSelect === 'both' ? 'שתי דלתות' : (sideSelect === 'right' ? 'ימין' : 'שמאל');
        readoutContent.innerHTML =
            `<div class="readout-item">סוג דלת: <strong>${doorType}</strong></div>
             <div class="readout-item">גובה חזית: <strong>${cabH} מ״מ</strong></div>
             <div class="readout-item">רוחב חזית: <strong>${frontW} מ״מ</strong></div>
             <div class="readout-item">צירים: <strong>${rEdge} / ${rMidStep.toFixed(0)} / ${rEdge} מ״מ</strong></div>`;
        const readout = document.getElementById('readout');
        if (readout) readout.style.display = 'block';
    }
}

//Listeners
sapakSelect.addEventListener("change", fillProfileOptions);
profileSelect.addEventListener("change", draw);
sideSelect.addEventListener("change", draw);
frontW.addEventListener("change", draw);
cabH.addEventListener("change", draw);
HingeLocation.addEventListener("change", draw);
HingeCount.addEventListener("change", draw);

// Load and process Excel file
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

        // המרה של GENESIS לפורמט MT59_ג'נסיס והעברת הקוד לעמודת מלואה
        excelRows = excelRows.map(r => {
            if (r['סוג החומר'] && String(r['סוג החומר']).toUpperCase().includes("GENESIS")) {
                const parts = String(r['סוג החומר']).split('_');

                // קוד זכוכית (כל מה אחרי _) אם קיים
                const glassCode = parts.length > 1 ? parts[1] : null;

                // סוג חומר בפורמט MT59_ג'נסיס
                r['סוג החומר'] = "MT59_ג'נסיס";

                // שמירת הקוד בעמודת מלואה
                if (glassCode) {
                    r['מלואה'] = glassCode;
                }
            }
            return r;
        });

        console.log("עמודות שהתקבלו:", Object.keys(excelRows[0]));
        console.log("דוגמה לשורה ראשונה:", excelRows[0]);

        // מציאת יחידות שיש להן לפחות דלת ימין/שמאל ומיון מספרי
        const validUnits = [...new Set(
            excelRows
                .filter(r => {
                    const partName = (r['שם החלק'] || "").toLowerCase();
                    return partName.includes("ימין") || partName.includes("שמאל");
                })
                .map(r => String(r['יחידה']).trim())
                .filter(u => u && u !== "undefined")
        )].sort((a, b) => Number(a) - Number(b));  // מיון מספרי

        if (validUnits.length === 0) {
            alert("לא נמצאה אף יחידה עם דלת ימין או שמאל בקובץ. לא ניתן להמשיך.");

            // השבתת כפתורים
            batchSaveBtn.disabled = true;
            batchSaveBtn.style.backgroundColor = "#ccc";
            batchSaveBtn.style.cursor = "not-allowed";

            downloadBtn.disabled = true;
            downloadBtn.style.backgroundColor = "#ccc";
            downloadBtn.style.cursor = "not-allowed";

            // איפוס השרטוט
            const svg = document.getElementById('svg');
            const overlay = document.querySelector('.svg-overlay');
            if (svg) svg.innerHTML = "";   // מוחק את תוכן ה־SVG
            if (overlay) overlay.style.display = 'none';

            return;
        } else {
            // הפעלה מחדש של כפתורים אם הכל תקין
            batchSaveBtn.disabled = false;
            batchSaveBtn.style.backgroundColor = "";
            batchSaveBtn.style.cursor = "pointer";

            downloadBtn.disabled = false;
            downloadBtn.style.backgroundColor = "";
            downloadBtn.style.cursor = "pointer";
        }

        // הפיכת השדה unitNum לרשימה נפתחת אם הוא עדיין input
        if (unitNumInput.tagName.toLowerCase() === "input") {
            const select = document.createElement("select");
            select.id = "unitNum";

            validUnits.forEach((unit, index) => {
                const option = document.createElement("option");
                option.value = unit;
                option.textContent = unit;
                select.appendChild(option);

                if (index === 0) select.value = unit;
            });

            unitContainer.replaceChild(select, unitNumInput);
            unitNumInput = select;
        }

        unitNumInput.addEventListener("change", function () {
            searchUnit(this.value);
        });

        searchUnit(unitNumInput.value);
    };
    reader.readAsArrayBuffer(file);
});

// Search and display unit details when unit number is selected or typed
unitNumInput.addEventListener("input", function () {
    searchUnit(this.value);
});

// Single PDF generation for the currently selected unit
downloadBtn.addEventListener("click", async (e) => {
    e.preventDefault();
    e.stopPropagation();
    try { await downloadPdf(); }
    catch (err) {
        console.error('[downloadPdf] failed:', err);
        alert('אירעה שגיאה בייצוא PDF. ראה קונסול.');
    }
});

// Batch generate PDFs for all corner cabinet units in the loaded Excel file
batchSaveBtn.addEventListener("click", async function () {
    if (!excelRows.length) return alert("אין קובץ Excel טעון!");

    const requiredFields = ['Sapak', 'planNum', 'unitNum', 'partName', 'profileType', 'profileColor', 'glassModel',];
    if (!validateRequiredFields(requiredFields)) return;

    showOverlay(); // מציג חלון המתנה

    // יצירת PDF לכל יחידה עם small delay כדי לייבא ערכים ל-DOM
    for (const row of excelRows) {
        if (!row['יחידה']) continue;

        const partName = row['שם החלק'] || '';

        // סינון: הורדה רק עבור דלת שמאל או דלת ימין
        if (!partName.includes('דלת שמאל') && !partName.includes('דלת ימין')) continue;

        const unitNumber = row['יחידה'];
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

// Update displayed file name when a new Excel file is selected
excelFileInput.addEventListener('change', () => {
    if (excelFileInput.files.length > 0) {
        fileNameSpan.textContent = excelFileInput.files[0].name;
    } else {
        fileNameSpan.textContent = "לא נבחר קובץ";
    }
});

// Add loading state to buttons
const buttons = document.querySelectorAll('button');
buttons.forEach(button => {
    button.addEventListener('click', function () {
        this.classList.add('loading');
        setTimeout(() => {
            this.classList.remove('loading');
        }, 2000);
    });
});

// First draw
draw();