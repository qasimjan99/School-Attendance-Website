let data = {};
let selected = "";
let result = { present: [], absent: [], unknown: [] };

function esc(v) { return String(v ?? "").replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[m])); }
function setStatus(msg, type = "success") { document.getElementById("fileStatus").innerHTML = `<div class="alert alert-${type} mb-0">${msg}</div>`; }
function populateSections() {
  const s = document.getElementById("sectionSelect");
  const names = Object.keys(data);
  s.innerHTML = names.map(x => `<option value="${esc(x)}">${esc(x)} (${data[x].length} students)</option>`).join("");
  if (names.length) { selected = names[0]; s.value = selected; selectSection(); }
}
function selectSection() {
  selected = document.getElementById("sectionSelect").value;
  const students = data[selected] || [];

  document.getElementById("total").textContent = students.length;
  document.getElementById("present").textContent = 0;
  document.getElementById("absent").textContent = 0;
  document.getElementById("unknown").textContent = 0;

  result = { present: [], absent: [], unknown: [] };

  renderTable("presentTable", [], "Present");
  renderTable("absentTable", [], "Absent");

  document.getElementById("presentTitle").textContent = selected || "-";
  document.getElementById("absentTitle").textContent = selected || "-";
  document.getElementById("warnings").innerHTML = "";
}
function normalizeHeader(x) { return String(x ?? "").toLowerCase().replace(/[.\\s_-]/g, ""); }
function loadExcel() {
  const f = document.getElementById("excelFile").files[0];
  if (!f) { setStatus("Please select an Excel file first.", "danger"); return; }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const parsed = {};
      wb.SheetNames.forEach(name => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "" });
        if (!rows.length) return;
        const h = rows[0].map(normalizeHeader);
        const idI = h.findIndex(x => x === "id" || x === "idno" || x === "studentid");
        const nameI = h.findIndex(x => x === "name" || x === "studentname");
        const fatherI = h.findIndex(x => x.includes("father") || x === "fname");
        const snoI = h.findIndex(x => x === "sno" || x === "srno" || x === "serialno");
        if (idI < 0) { return; }
        parsed[name] = rows.slice(1).map((r, i) => ({
          sno: r[snoI >= 0 ? snoI : 0] || i + 1,
          id: String(r[idI] ?? "").trim(),
          name: String(r[nameI >= 0 ? nameI : 2] ?? "").trim(),
          father: String(r[fatherI >= 0 ? fatherI : 3] ?? "").trim()
        })).filter(x => x.id);
      });
      if (!Object.keys(parsed).length) { setStatus("No valid student sheets found. Make sure each sheet has S.No, I.D, Name and F/Name columns.", "danger"); return; }
      data = parsed; populateSections();
      setStatus(`Loaded ${Object.keys(data).length} sections from ${f.name}.`, "success");
    } catch (err) { setStatus("Could not read the Excel file: " + err.message, "danger"); }
  };
  reader.readAsArrayBuffer(f);
}
function updateStats(total, present, absent, unknown) {
  document.getElementById('total').innerText = total;
  document.getElementById('present').innerText = present;
  document.getElementById('absent').innerText = absent;
  document.getElementById('unknown').innerText = unknown;


  
  result = { present: [], absent: [], unknown: [] };
  renderTable("presentTable", [], "Present"); renderTable("absentTable", [], "Absent");
  document.getElementById("presentTitle").textContent = selected || "-";
  document.getElementById("absentTitle").textContent = selected || "-";
  document.getElementById("warnings").innerHTML = "";
}
function generateAttendance() {
  if (!selected) { alert("Select a class/section first."); return; }
  const students = data[selected] || [];
  const snos = [...new Set(document.getElementById("snoInput").value.split(/[\s,;]+/).map(x => x.trim()).filter(Boolean))];
  const map = new Map(students.map(s => [String(s.sno), s]));
  result.unknown = snos.filter(sno => !map.has(sno));
  const absentSet = new Set(snos.filter(sno => map.has(sno)));
  result.absent = students.filter(s => absentSet.has(String(s.sno)));
  result.present = students.filter(s => !absentSet.has(String(s.sno)));
  document.getElementById("total").textContent = students.length;
  document.getElementById("present").textContent = result.present.length;
  document.getElementById("absent").textContent = result.absent.length;
  document.getElementById("unknown").textContent = result.unknown.length;
  renderTable("presentTable", result.present, "Present");
  renderTable("absentTable", result.absent, "Absent");
  document.getElementById("warnings").innerHTML = result.unknown.length
    ? `<div class="alert alert-warning"><b>Unknown S.No:</b> ${esc(result.unknown.join(", "))}</div>`
    : `<div class="alert alert-success">Attendance generated successfully for <b>${esc(selected)}</b>.</div>`;
}
function renderTable(id, rows, status) {
  const t = document.getElementById(id);
  if (!rows.length) { t.innerHTML = '<tbody><tr><td class="text-secondary">No data yet.</td></tr></tbody>'; return; }
  t.innerHTML = `<thead class="table-light sticky-top"><tr><th>S.No</th><th>ID</th><th>Name</th><th>Father Name</th><th>Status</th></tr></thead><tbody>` +
    rows.map(s => `<tr><td>${esc(s.sno)}</td><td>${esc(s.id)}</td><td>${esc(s.name)}</td><td>${esc(s.father)}</td><td><span class="badge ${status === "Present" ? "text-bg-success" : "text-bg-danger"}">${status}</span></td></tr>`).join("") + "</tbody>";
}
function exportSection(type) {
  if (!selected || !result[type].length) { alert("Generate attendance first."); return; }
  const label = type === "present" ? "Present" : "Absent";
  const ws = createAttendanceSheet(result[type], label);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, type === "present" ? "Present" : "Absent");
  XLSX.writeFile(wb, `${safeName(selected)}_${type === "present" ? "Present" : "Absent"}.xlsx`);
}
function exportComplete() {
  if (!selected || (!result.present.length && !result.absent.length)) { alert("Generate attendance first."); return; }
  const wb = XLSX.utils.book_new();
  const p = createAttendanceSheet(result.present, "Present");
  const a = createAttendanceSheet(result.absent, "Absent");
  XLSX.utils.book_append_sheet(wb, p, "Present"); XLSX.utils.book_append_sheet(wb, a, "Absent");
  XLSX.writeFile(wb, `${safeName(selected)}_Attendance.xlsx`);
}
function createAttendanceSheet(students, status) {
  const minimumRows = students.length;
  const today = new Date().toLocaleDateString("en-GB");
  const rows = [
    ["JAMAL INTERNATIONAL DEGREE COLLEGE SCHOOL SYSTEM", "", "", "", ""],
    ["Dora Road Kohat Road Peshawar", "", "", "", ""],
    [`${status} Students Class ${selected}`, "", "", "", today],
    ["", "", "", "", ""],
    ["S.No", "I.D", "Name", "F/Name", status]
  ];
  for (let index = 0; index < minimumRows; index++) {
    const student = students[index];
    rows.push([index + 1, student?.id || "", student?.name || "", student?.father || "", student ? status : ""]);
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } }
  ];
  ws["!cols"] = [{ wch: 9 }, { wch: 13 }, { wch: 34 }, { wch: 34 }, { wch: 14 }];
  const border = { top: { style: "thin", color: { rgb: "000000" } }, bottom: { style: "thin", color: { rgb: "000000" } }, left: { style: "thin", color: { rgb: "000000" } }, right: { style: "thin", color: { rgb: "000000" } } };
  const titleStyle = { font: { name: "Times New Roman", sz: 16, bold: true, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" } };
  const subtitleStyle = { font: { name: "Times New Roman", sz: 14, bold: true, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" } };
  const headerStyle = { font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "000000" } }, fill: { fgColor: { rgb: "FFC000" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border };
  ws["A1"].s = titleStyle; ws["A2"].s = subtitleStyle; ws["A3"].s = subtitleStyle;
  ws["E3"].s = { ...subtitleStyle, alignment: { horizontal: "right", vertical: "center" } };
  for (let col = 0; col < 5; col++) {
    ws[XLSX.utils.encode_cell({ r: 4, c: col })].s = headerStyle;
  }
  for (let row = 5; row < rows.length; row++) {
    for (let col = 0; col < 5; col++) {
      const cell = XLSX.utils.encode_cell({ r: row, c: col });
      if (!ws[cell]) ws[cell] = { v: "", t: "s" };
      ws[cell].s = { border, alignment: { horizontal: col === 0 || col === 1 || col === 4 ? "center" : "left", vertical: "center" } };
    }
  }
  ws["!rows"] = [{ hpt: 26 }, { hpt: 23 }, { hpt: 24 }, { hpt: 8 }, { hpt: 25 }];
  for (let row = 5; row < rows.length; row++)ws["!rows"][row] = { hpt: 20 };
  ws["!freeze"] = { xSplit: 0, ySplit: 5 };
  ws["!autofilter"] = { ref: `A5:E${rows.length}` };
  ws["!pageSetup"] = { orientation: "portrait", paperSize: 9, fitToWidth: 1, fitToHeight: 0, scale: 100 };
  ws["!pageMargins"] = { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
  ws["!printOptions"] = { horizontalCentered: true, verticalCentered: false };
  ws["!printTitles"] = { rows: "1:5" };
  return ws;
}
// Complete Excel Export Function (Styled Header & Mixed Sequence)
function exportComplete() {
  if (!selected || (!result.present.length && !result.absent.length)) {
    alert("Generate attendance first.");
    return;
  }

  // 1. Present aur Absent ko mix karke Serial Number (id/sno) ke mutabiq sort karna
  const combinedData = [
    ...result.present.map(s => ({ ...s, status: "Present" })),
    ...result.absent.map(s => ({ ...s, status: "Absent" }))
  ].sort((a, b) => (parseInt(a.sno || a.id) || 0) - (parseInt(b.sno || b.id) || 0));

  // 2. Custom Rows Header and Data preparation
  const today = new Date().toLocaleDateString("en-GB");
  const rows = [
    ["JAMAL INTERNATIONAL DEGREE COLLEGE SCHOOL SYSTEM", "", "", "", ""],
    ["Dora Road Kohat Road Peshawar", "", "", "", ""],
    [`Complete Attendance Report - Class ${selected}`, "", "", "", `Date: ${today}`],
    ["", "", "", "", ""], // Blank separator row
    ["S.No", "I.D", "Name", "Father Name", "Status"]
  ];

  combinedData.forEach((student, index) => {
    rows.push([
      index + 1,
      student.id || student.sno || "",
      student.name || "",
      student.father || "",
      student.status
    ]);
  });

  // 3. Create Sheet with Styling
  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Merges & Column Widths
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } }
  ];
  ws["!cols"] = [{ wch: 9 }, { wch: 13 }, { wch: 34 }, { wch: 34 }, { wch: 14 }];

  // Styling Rules
  const border = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  };
  const titleStyle = { font: { name: "Times New Roman", sz: 16, bold: true, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" } };
  const subtitleStyle = { font: { name: "Times New Roman", sz: 14, bold: true, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" } };
  const headerStyle = { font: { name: "Calibri", sz: 12, bold: true, color: { rgb: "000000" } }, fill: { fgColor: { rgb: "FFC000" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border };

  if (ws["A1"]) ws["A1"].s = titleStyle;
  if (ws["A2"]) ws["A2"].s = subtitleStyle;
  if (ws["A3"]) ws["A3"].s = subtitleStyle;
  if (ws["E3"]) ws["E3"].s = { ...subtitleStyle, alignment: { horizontal: "right", vertical: "center" } };

  // Apply Table Header Style
  for (let col = 0; col < 5; col++) {
    const cell = XLSX.utils.encode_cell({ r: 4, c: col });
    if (ws[cell]) ws[cell].s = headerStyle;
  }

  // Apply Data Rows Style & Borders
  for (let row = 5; row < rows.length; row++) {
    for (let col = 0; col < 5; col++) {
      const cell = XLSX.utils.encode_cell({ r: row, c: col });
      if (!ws[cell]) ws[cell] = { v: "", t: "s" };
      ws[cell].s = {
        border,
        alignment: { horizontal: col === 0 || col === 1 || col === 4 ? "center" : "left", vertical: "center" }
      };
    }
  }

  // Row Heights & Print Settings
  ws["!rows"] = [{ hpt: 26 }, { hpt: 23 }, { hpt: 24 }, { hpt: 8 }, { hpt: 25 }];
  for (let row = 5; row < rows.length; row++) ws["!rows"][row] = { hpt: 20 };

  ws["!freeze"] = { xSplit: 0, ySplit: 5 };
  ws["!autofilter"] = { ref: `A5:E${rows.length}` };
  ws["!pageSetup"] = { orientation: "portrait", paperSize: 9, fitToWidth: 1, fitToHeight: 0, scale: 100 };
  ws["!pageMargins"] = { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
  ws["!printOptions"] = { horizontalCentered: true, verticalCentered: false };
  ws["!printTitles"] = { rows: "1:5" };

  // Export File
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Complete Attendance");
  XLSX.writeFile(wb, `${safeName(selected)}_Complete_Attendance.xlsx`);
}

// PDF Download Function
function downloadPDF(type) {
  if (!selected || (!result.present.length && !result.absent.length)) { 
    alert("Generate attendance first."); 
    return; 
  }
  if (!window.jspdf || typeof window.jspdf.jsPDF !== "function") { 
    alert("PDF library could not be loaded. Check your internet connection and try again."); 
    return; 
  }

  let reports = [];

  if (type === "complete") {
    const combinedStudents = [
      ...result.present.map(s => ({ ...s, status: "Present" })),
      ...result.absent.map(s => ({ ...s, status: "Absent" }))
    ].sort((a, b) => (parseInt(a.sno || a.id) || 0) - (parseInt(b.sno || b.id) || 0));

    reports = [{ students: combinedStudents, status: "Complete Attendance" }];
  } else {
    reports = [{ 
      students: result[type].map(s => ({ ...s, status: type === "present" ? "Present" : "Absent" })), 
      status: type === "present" ? "Present" : "Absent" 
    }];
  }

  const printableReports = reports.filter(report => report.students.length);
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const today = new Date().toLocaleDateString("en-GB");

  printableReports.forEach((report, index) => {
    if (index) pdf.addPage();
    pdf.setFont("times", "bold"); 
    pdf.setFontSize(16); 
    pdf.text("JAMAL INTERNATIONAL DEGREE COLLEGE SCHOOL SYSTEM", 105, 16, { align: "center" });
    
    pdf.setFontSize(14); 
    pdf.text("Dora Road Kohat Road Peshawar", 105, 24, { align: "center" });
    pdf.text(`${report.status} Class ${selected}`, 14, 34); 
    pdf.text(today, 196, 34, { align: "right" });

    pdf.autoTable({ 
      startY: 39, 
      head: [["S.No", "I.D", "Name", "F/Name", "Status"]], 
      body: report.students.map((student, idx) => [
        idx + 1, 
        student.id || student.sno || "", 
        student.name || "", 
        student.father || "", 
        student.status
      ]), 
      theme: "grid", 
      styles: { font: "helvetica", fontSize: 9, textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.2, cellPadding: 2.2 }, 
      headStyles: { fillColor: [255, 192, 0], textColor: [0, 0, 0], fontStyle: "bold", halign: "center" }, 
      columnStyles: { 
        0: { halign: "center", cellWidth: 13 }, 
        1: { halign: "center", cellWidth: 18 }, 
        2: { cellWidth: 57 }, 
        3: { cellWidth: 57 }, 
        4: { halign: "center", cellWidth: 25 } 
      } 
    });
  });

  pdf.save(`${safeName(selected)}_${type === "complete" ? "Complete_Attendance" : type}.pdf`);
}

function safeName(s) { return s.replace(/[\\\\/:*?"<>|]/g, "-"); }

function clearEntry() { 
  document.getElementById("snoInput").value = ""; 
  document.getElementById("warnings").innerHTML = ""; 
}

function resetAll() { 
  data = {}; 
  selected = ""; 
  result = { present: [], absent: [], unknown: [] }; 
  document.getElementById("sectionSelect").innerHTML = ""; 
  clearEntry(); 
  document.getElementById("fileStatus").innerHTML = ""; 
  document.getElementById("total").textContent = 0; 
  document.getElementById("present").textContent = 0; 
  document.getElementById("absent").textContent = 0; 
  document.getElementById("unknown").textContent = 0; 
  renderTable("presentTable", [], "Present"); 
  renderTable("absentTable", [], "Absent"); 
}






