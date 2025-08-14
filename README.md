# 📄 Download Word File from HTML Report (PL/SQL Dynamic Content) in Oracle APEX using JavaScript Library

This guide helps you integrate a **"Download Word"** feature in your Oracle APEX application using html2pdf.js.

---

## 📁 1. Upload the JavaScript Library

1. Download the html2pdf.bundle.min.js library:
   👉 https://ekoopmans.github.io/html2pdf.js/

3. In Oracle APEX:
   - Go to **Shared Components** → **Static Application Files**
   - Click **Upload File**
   - Upload the file html2pdf.bundle.min.js
   - #APP_FILES#html2pdf.bundle.min.js

---
## 🧠 2. Add JavaScript to Page-Level Global Variable Function

 - Add This Static File to Page HTML Header.
```HTML Header
<script src="#APP_FILES#html2pdf.bundle.min.js"></script>
```

## 🧠 3. Add JavaScript to Page-Level Global Variable Function

1. Go to the page where you want the export feature.
2. Under **Page Attributes**, scroll to **Function and Global Variable Declaration**.
3. Paste the following code:

```javascript

function exportHTMLToWord(elementId, filename) {
  var header = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' 
          xmlns:w='urn:schemas-microsoft-com:office:word' 
          xmlns='http://www.w3.org/TR/REC-html40'>
    <head><meta charset='utf-8'></head><body>`;
  var footer = "</body></html>";

  var content = document.getElementById(elementId).innerHTML;
  var html = header + content + footer;

  var blob = new Blob(['\ufeff', html], {
    type: 'application/msword'
  });

  // Default file name
  var docName = filename || 'Invoice.doc';

  // Create download link
  var link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = docName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}


```

## 4. 📥 Download Word via Dynamic Action (Button)

### ✅ Step-by-Step:

1. **Create a Button**
   - Label it (e.g., `Download Word`)
   - Position it on the same APEX page as your report (HTML container with `printID`).

2. **Create a Dynamic Action**
   - **Event**: `Click`
   - **Selection Type**: `Button`
   - **Button**: Select the button you just created (e.g., `Download Word`)

3. **Add a True Action**
   - **Action Type**: `Execute JavaScript Code`
   - **Code**:
     ```javascript
     exportHTMLToWord('printID', 'Invoice-' + new Date().toISOString().slice(0,10) + '.doc');

     ```

### 💡 Notes:
- Ensure the element you want to print has the ID `printID`.
- `html2pdf.js` must be loaded on the page (see section 2 of the documentation for CDN or static file setup).
- The `exportHTMLToWord` function must be declared globally (in Page > JavaScript > Function and Global Variable Declaration).
- Create a Dynamic Action for the button’s click event to call the JavaScript function.

---


# Thank You
## Sanjay Sikder
