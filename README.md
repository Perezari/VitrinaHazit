# 📐 Vitrina Hazit Configurator

![Language](https://img.shields.io/badge/language-JavaScript-yellow.svg)
![HTML/CSS](https://img.shields.io/badge/markup_styling-HTML_CSS-orange.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## Description

This project is a web-based application designed to streamline the process of generating technical drawings and PDF documents for custom vitrine/cabinet fronts. Users can input precise dimensions and material specifications (such as profile type, color, and glass model), and the system will automatically render a detailed SVG technical drawing. It also supports batch processing, allowing users to upload an Excel file to generate multiple PDF drawings efficiently.

The tool features dynamic input fields, supplier-specific configurations, and direct PDF generation, making it an invaluable asset for professionals in the carpentry, kitchen, or furniture design industries.

✨ **Features**

*   **Dynamic SVG Drawing:** Generates interactive and scalable technical drawings of vitrine fronts in real-time based on user inputs.
*   **Customizable Dimensions:** Easily define front width, height, hinge locations, and hinge count.
*   **Detailed Specifications:** Input supplier, plan number, unit number, part name, profile type, profile color, glass model, and glass texture.
*   **Supplier & Profile Management:** Dynamically loads available profile types based on the selected supplier, utilizing a comprehensive `ProfileConfig`.
*   **Editable SVG Dimensions:** Users can directly double-click and edit dimension values on the SVG drawing for fine-tuning.
*   **Single PDF Export:** Download a single technical drawing as a PDF document with all specified details and a dynamic supplier logo.
*   **Batch PDF Generation:** Upload an Excel file (`.xls`, `.xlsx`) to automatically generate and download multiple PDFs for various units, extracting data like dimensions and material types.
*   **Excel Data Integration:** Automatically parses Excel files to pre-fill form fields and identify units requiring drawings.
*   **Hebrew Language & RTL Support:** Fully localized for Hebrew speakers with right-to-left layout and specialized font handling for PDF output.
*   **Configurable Profiles:** Supports various profile types (e.g., "ג'נסיס", "קואדרו", "זירו") with distinct drawing rules and drill hole specifications.

📚 **Tech Stack**

*   **HTML5:** For structuring the web content.
*   **CSS3:** For styling and responsive design (vanilla CSS with Google Fonts - Rubik).
*   **JavaScript:** Core logic for dynamic rendering, user interaction, and data processing (vanilla JS).
*   **jsPDF:** JavaScript library for generating PDFs client-side.
*   **svg2pdf.js:** A plugin for jsPDF to convert SVG elements into PDF documents.
*   **XLSX.js:** Library for reading and parsing Excel files.
*   **Alef-Normal.js:** Custom font for ensuring correct Hebrew rendering in generated PDFs.

🚀 **Installation**

To get this project up and running locally, follow these simple steps:

1.  **Clone the repository:**

    ```bash
    git clone https://github.com/your-username/VitrinaHazit.git
    ```

2.  **Navigate to the project directory:**

    ```bash
    cd VitrinaHazit
    ```

3.  **Open the `index.html` file:**
    Since this is a client-side application, you can simply open the `index.html` file directly in your web browser.

    ```bash
    # On macOS
    open index.html
    # On Windows
    start index.html
    ```

    Alternatively, you can serve it with a local web server (e.g., using Python's `http.server` or `Live Server` VS Code extension) to ensure all features, especially Excel file uploads, work correctly without browser security restrictions.

    ```bash
    # Using Python (from the project root)
    python -m http.server 8000
    ```
    Then, open your browser and navigate to `http://localhost:8000/`.

▶️ **Usage**

Once the application is loaded in your browser:

1.  **Input Dimensions:** Enter the `רוחב חזית` (Front Width), `גובה חזית` (Front Height), `מיקום ציר תחתון ועליון` (Bottom and Top Hinge Location), and `כמות צירים דרושה` (Required Hinge Count) in millimeters. The SVG drawing will update dynamically.
2.  **Specify Unit Details:** Fill in the `פרטי יחידה` (Unit Details) section, including `הזמנה עבור` (Supplier), `מספר תוכנית` (Plan Number), `מספר יחידה` (Unit Number), `שם מפרק` (Part Name), `סוג פרופיל` (Profile Type), `גוון פרופיל` (Profile Color), `דגם זכוכית` (Glass Model), `כיוון טקסטורת זכוכית` (Glass Texture), and `כולל הכנה עבור` (Prepared for).
3.  **Download Single PDF:** Click the `הורד PDF 💾` button to generate and download a PDF for the current unit's specifications.
4.  **Batch PDF Generation (from Excel):**
    *   Click the `...` button next to `לא נבחר קובץ` (No file selected) and upload an Excel file (`.xls` or `.xlsx`) containing your unit data.
    *   The `מספר תוכנית` (Plan Number) will be automatically extracted from the file name.
    *   The `מספר יחידה` (Unit Number) field will turn into a dropdown, allowing you to select specific units from the Excel data. Selecting a unit will auto-fill the other fields.
    *   Click the `PDF BATCH 💾` button to generate PDFs for all valid units found in the Excel file. An overlay will indicate the saving process.

🤝 **Contributing**

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

📝 **License**

This project is distributed under the MIT License. See the [LICENSE](LICENSE) file for more information. (Note: A LICENSE file was not found in the provided data. It is recommended to create one if you intend to share this project.)
