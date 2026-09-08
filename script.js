let data = {};
let pendingData = {};
let selected = "";
let result = {present:[], absent:[], unknown:[]};

function esc(v){return String(v ?? "").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]));}
function setStatus(msg,type="success"){document.getElementById("fileStatus").innerHTML=`<div class="alert alert-${type} mb-0">${msg}</div>`;}
function populateSections(){
  const s=document.getElementById("sectionSelect");
  const names=Object.keys(data);
  s.innerHTML=names.map(x=>`<option value="${esc(x)}">${esc(x)} (${data[x].length} students)</option>`).join("");
  if(names.length){selected=names[0];s.value=selected;selectSection();}
}
function useBuiltIn(){
  const oldMasterDataCleared = "schoolAttendanceMasterDataCleared";
  if(!localStorage.getItem(oldMasterDataCleared)){
    clearSharedAttendanceData();
    localStorage.setItem(oldMasterDataCleared,"true");
  }
  data=loadSharedAttendanceData() || {};
  if(Object.keys(data).length){
    populateSections();
    setStatus(`Loaded ${Object.keys(data).length} saved sections.`,"success");
  }else{
    document.getElementById("sectionSelect").innerHTML="";
    setStatus("No Excel data loaded. Please upload your Excel files.","info");
  }
}
function normalizeHeader(x){return String(x??"").toLowerCase().replace(/[.\\s_-]/g,"");}
function readWorkbook(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{resolve(XLSX.read(e.target.result,{type:"array"}));}
      catch(error){reject(error);}
    };
    reader.onerror=()=>reject(new Error(`Could not read ${file.name}.`));
    reader.readAsArrayBuffer(file);
  });
}
async function loadExcel(){
  const files=[...document.getElementById("excelFile").files];
  if(!files.length){setStatus("Please select one or more Excel files first.","danger");return;}
  try{
    const parsed={};
    for(const file of files){
      const wb=await readWorkbook(file);
      wb.SheetNames.forEach(name=>{
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:""});
        if(!rows.length)return;
        const h=rows[0].map(normalizeHeader);
        const idI=h.findIndex(x=>x==="id"||x==="idno"||x==="studentid");
        const nameI=h.findIndex(x=>x==="name"||x==="studentname");
        const fatherI=h.findIndex(x=>x.includes("father")||x==="fname");
        const snoI=h.findIndex(x=>x==="sno"||x==="srno"||x==="serialno");
        if(idI<0)return;
        parsed[name]=rows.slice(1).map((r,i)=>({
          sno:r[snoI>=0?snoI:0]||i+1,
          id:String(r[idI]??"").trim(),
          name:String(r[nameI>=0?nameI:2]??"").trim(),
          father:String(r[fatherI>=0?fatherI:3]??"").trim()
        })).filter(x=>x.id);
      });
    }
    if(!Object.keys(parsed).length){setStatus("No valid student sheets found. Make sure each sheet has S.No, I.D, Name and F/Name columns.","danger");return;}
    pendingData=parsed;
    data={...(loadSharedAttendanceData() || {}),...pendingData};
    populateSections();
    setStatus(`Loaded ${Object.keys(parsed).length} sections. Press Save Data to keep them.`,"info");
  }catch(error){setStatus("Could not read the Excel file: "+error.message,"danger");}
}
function saveData(){
  if(!Object.keys(pendingData).length){
    setStatus("Please load one or more Excel files first.","danger");
    return;
  }
  const savedData=loadSharedAttendanceData() || {};
  data={...savedData,...pendingData};
  saveSharedAttendanceData(data);
  pendingData={};
  populateSections();
  setStatus(`Saved ${Object.keys(data).length} sections. Previous saved data was kept.`,"success");
}
function removeData(){
  data={};
  pendingData={};
  selected="";
  clearSharedAttendanceData();
  document.getElementById("excelFile").value="";
  document.getElementById("sectionSelect").innerHTML="";
  clearEntry();
  document.getElementById("total").textContent=0;
  document.getElementById("present").textContent=0;
  document.getElementById("absent").textContent=0;
  document.getElementById("unknown").textContent=0;
  renderTable("presentTable",[],"Present");
  renderTable("absentTable",[],"Absent");
  setStatus("Uploaded Excel master data removed.","warning");
}
function selectSection(){
  selected=document.getElementById("sectionSelect").value;
  const list=data[selected]||[];
  document.getElementById("total").textContent=list.length;
  document.getElementById("present").textContent=0;
  document.getElementById("absent").textContent=0;
  document.getElementById("unknown").textContent=0;
  result={present:[],absent:[],unknown:[]};
  renderTable("presentTable",[],"Present");renderTable("absentTable",[],"Absent");
  document.getElementById("presentTitle").textContent=selected||"-";
  document.getElementById("absentTitle").textContent=selected||"-";
  document.getElementById("warnings").innerHTML="";
}
function generateAttendance(){
  if(!selected){alert("Select a class/section first.");return;}
  const attendanceDate=document.getElementById("attendanceDate").value;
  if(!attendanceDate){alert("Select the attendance date first.");return;}
  const students=data[selected]||[];
  const snos=[...new Set(document.getElementById("snoInput").value.split(/[\s,;]+/).map(x=>x.trim()).filter(Boolean))];
  const map=new Map(students.map(s=>[String(s.sno),s]));
  result.unknown=snos.filter(sno=>!map.has(sno));
  const absentSet=new Set(snos.filter(sno=>map.has(sno)));
  result.absent=students.filter(s=>absentSet.has(String(s.sno)));
  result.present=students.filter(s=>!absentSet.has(String(s.sno)));
  document.getElementById("total").textContent=students.length;
  document.getElementById("present").textContent=result.present.length;
  document.getElementById("absent").textContent=result.absent.length;
  document.getElementById("unknown").textContent=result.unknown.length;
  renderTable("presentTable",result.present,"Present");
  renderTable("absentTable",result.absent,"Absent");
    saveDailyAttendance(selected,attendanceDate,{
      students:students.map(student=>({...student})),
      present:result.present.map(student=>String(student.sno)),
      absent:result.absent.map(student=>String(student.sno)),
      unknown:result.unknown
    });
  document.getElementById("warnings").innerHTML=result.unknown.length
    ? `<div class="alert alert-warning"><b>Unknown S.No:</b> ${esc(result.unknown.join(", "))}</div>`
    : `<div class="alert alert-success">Attendance generated successfully for <b>${esc(selected)}</b>.</div>`;
}
function renderTable(id,rows,status){
  const t=document.getElementById(id);
  if(!rows.length){t.innerHTML='<tbody><tr><td class="text-secondary">No data yet.</td></tr></tbody>';return;}
  t.innerHTML=`<thead class="table-light sticky-top"><tr><th>S.No</th><th>ID</th><th>Name</th><th>Father Name</th><th>Status</th></tr></thead><tbody>`+
    rows.map(s=>`<tr><td>${esc(s.sno)}</td><td>${esc(s.id)}</td><td>${esc(s.name)}</td><td>${esc(s.father)}</td><td><span class="badge ${status==="Present"?"text-bg-success":"text-bg-danger"}">${status}</span></td></tr>`).join("")+"</tbody>";
}
function exportSection(type){
  if(!selected||!result[type].length){alert("Generate attendance first.");return;}
  const label=type==="present"?"Present":"Absent";
  const ws=createAttendanceSheet(result[type],label);
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,type==="present"?"Present":"Absent");
  XLSX.writeFile(wb,`${safeName(selected)}_${type==="present"?"Present":"Absent"}.xlsx`);
}
function exportComplete(){
  if(!selected||(!result.present.length&&!result.absent.length)){alert("Generate attendance first.");return;}
  const wb=XLSX.utils.book_new();
  const statusBySno=new Map(result.absent.map(student=>[String(student.sno),"Absent"]));
  result.present.forEach(student=>statusBySno.set(String(student.sno),"Present"));
  const completeStudents=(data[selected]||[]).map(student=>({
    ...student,
    attendanceStatus:statusBySno.get(String(student.sno)) || ""
  }));
  const completeSheet=createAttendanceSheet(completeStudents,"Complete Attendance",true);
  XLSX.utils.book_append_sheet(wb,completeSheet,"Complete");
  XLSX.writeFile(wb,`${safeName(selected)}_Attendance.xlsx`);
}
function createAttendanceSheet(students,status,highlightAbsent=false){
  const minimumRows=students.length;
  const today=new Date().toLocaleDateString("en-GB");
  const rows=[
    ["JAMAL INTERNATIONAL DEGREE COLLEGE SCHOOL SYSTEM","","","", ""],
    ["Dora Road Kohat Road Peshawar","","","", ""],
    [`${status} - Class ${selected}`,"","","",today],
    ["","","","", ""],
    ["S.No","I.D","Name","F/Name",status === "Complete Attendance" ? "Status" : status]
  ];
  for(let index=0;index<minimumRows;index++){
    const student=students[index];
    rows.push([index+1,student?.id||"",student?.name||"",student?.father||"",student ? (student.attendanceStatus || status) : ""]);
  }
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"]=[
    {s:{r:0,c:0},e:{r:0,c:4}},
    {s:{r:1,c:0},e:{r:1,c:4}},
    {s:{r:2,c:0},e:{r:2,c:3}}
  ];
  ws["!cols"]=[{wch:9},{wch:13},{wch:34},{wch:34},{wch:14}];
  const border={top:{style:"thin",color:{rgb:"000000"}},bottom:{style:"thin",color:{rgb:"000000"}},left:{style:"thin",color:{rgb:"000000"}},right:{style:"thin",color:{rgb:"000000"}}};
  const titleStyle={font:{name:"Times New Roman",sz:16,bold:true,color:{rgb:"000000"}},alignment:{horizontal:"center",vertical:"center"}};
  const subtitleStyle={font:{name:"Times New Roman",sz:14,bold:true,color:{rgb:"000000"}},alignment:{horizontal:"center",vertical:"center"}};
  const headerStyle={font:{name:"Calibri",sz:12,bold:true,color:{rgb:"000000"}},fill:{fgColor:{rgb:"FFC000"}},alignment:{horizontal:"center",vertical:"center",wrapText:true},border};
  ws["A1"].s=titleStyle;ws["A2"].s=subtitleStyle;ws["A3"].s=subtitleStyle;
  ws["E3"].s={...subtitleStyle,alignment:{horizontal:"right",vertical:"center"}};
  for(let col=0;col<5;col++){
    ws[XLSX.utils.encode_cell({r:4,c:col})].s=headerStyle;
  }
  for(let row=5;row<rows.length;row++){
    for(let col=0;col<5;col++){
      const cell=XLSX.utils.encode_cell({r:row,c:col});
      if(!ws[cell])ws[cell]={v:"",t:"s"};
      ws[cell].s={border,alignment:{horizontal:col===0||col===1||col===4?"center":"left",vertical:"center"}};
      if(highlightAbsent && col===4 && ws[cell].v==="Absent"){
        ws[cell].s.fill={fgColor:{rgb:"FFC7CE"}};
        ws[cell].s.font={bold:true,color:{rgb:"9C0006"}};
      }
    }
  }
  ws["!rows"]=[{hpt:26},{hpt:23},{hpt:24},{hpt:8},{hpt:25}];
  for(let row=5;row<rows.length;row++)ws["!rows"][row]={hpt:20};
  ws["!freeze"]={xSplit:0,ySplit:5};
  ws["!autofilter"]={ref:`A5:E${rows.length}`};
  ws["!pageSetup"]={orientation:"portrait",paperSize:9,fitToWidth:1,fitToHeight:0,scale:100};
  ws["!pageMargins"]={left:0.25,right:0.25,top:0.5,bottom:0.5,header:0.2,footer:0.2};
  ws["!printOptions"]={horizontalCentered:true,verticalCentered:false};
  ws["!printTitles"]={rows:"1:5"};
  return ws;
}
function downloadPDF(type){
  if(!selected||(!result.present.length&&!result.absent.length)){alert("Generate attendance first.");return;}
  if(!window.jspdf||typeof window.jspdf.jsPDF!=="function"){alert("PDF library could not be loaded. Check your internet connection and try again.");return;}
  const statusBySno=new Map(result.absent.map(student=>[String(student.sno),"Absent"]));
  result.present.forEach(student=>statusBySno.set(String(student.sno),"Present"));
  const completeStudents=(data[selected]||[]).map(student=>({...student,attendanceStatus:statusBySno.get(String(student.sno)) || ""}));
  const reports=type==="complete"
    ? [{students:completeStudents,status:"Complete",mixed:true}]
    : [{students:result[type],status:type==="present"?"Present":"Absent",mixed:false}];
  const printableReports=reports.filter(report=>report.students.length);
  const {jsPDF}=window.jspdf;
  const pdf=new jsPDF({orientation:"portrait",unit:"mm",format:"a4"});
  const today=new Date().toLocaleDateString("en-GB");
  printableReports.forEach((report,index)=>{
    if(index)pdf.addPage();
    pdf.setFont("times","bold");pdf.setFontSize(16);pdf.text("JAMAL INTERNATIONAL DEGREE COLLEGE SCHOOL SYSTEM",105,16,{align:"center"});
    pdf.setFontSize(14);pdf.text("Dora Road Kohat Road Peshawar",105,24,{align:"center"});
    pdf.setFontSize(14);pdf.text(`${report.status} Attendance Class ${selected}`,14,34);pdf.text(today,196,34,{align:"right"});
    const reportRows=report.students.map((student,index)=>[index+1,student.id,student.name,student.father,report.mixed ? student.attendanceStatus : report.status]);
    pdf.autoTable({startY:39,head:[["S.No","I.D","Name","F/Name",report.mixed?"Status":report.status]],body:reportRows,theme:"grid",styles:{font:"helvetica",fontSize:9,textColor:[0,0,0],lineColor:[0,0,0],lineWidth:0.2,cellPadding:2.2},headStyles:{fillColor:[255,192,0],textColor:[0,0,0],fontStyle:"bold",halign:"center"},columnStyles:{0:{halign:"center",cellWidth:13},1:{halign:"center",cellWidth:18},2:{cellWidth:57},3:{cellWidth:57},4:{halign:"center",cellWidth:25}},didParseCell:hookData=>{
      if(report.mixed && hookData.section==="body" && hookData.column.index===4 && hookData.cell.raw==="Absent"){
        hookData.cell.styles.fillColor=[220,53,69];
        hookData.cell.styles.textColor=[255,255,255];
        hookData.cell.styles.fontStyle="bold";
      }
    }});
  });
  pdf.save(`${safeName(selected)}_${type==="complete"?"Attendance":type}.pdf`);
}
function safeName(s){return s.replace(/[\\\\/:*?"<>|]/g,"-");}
function clearEntry(){document.getElementById("snoInput").value="";document.getElementById("warnings").innerHTML="";}
function resetAll(){
  data={};pendingData={};selected="";result={present:[],absent:[],unknown:[]};
  clearSharedAttendanceData();
  clearDailyAttendance();
  document.getElementById("excelFile").value="";
  document.getElementById("sectionSelect").innerHTML="";
  clearEntry();
  document.getElementById("fileStatus").innerHTML="";
  document.getElementById("total").textContent=0;
  document.getElementById("present").textContent=0;
  document.getElementById("absent").textContent=0;
  document.getElementById("unknown").textContent=0;
  renderTable("presentTable",[],"Present");
  renderTable("absentTable",[],"Absent");
}
function setToday(){
  const input=document.getElementById("attendanceDate");
  if(input&&!input.value)input.value=new Date().toISOString().slice(0,10);
}
useBuiltIn();
setToday();
