const PYODIDE_URL="https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";
const $=id=>document.getElementById(id);
let pyodide=null;
let engineReady=false;

function setStatus(text,type=""){
  const node=$("engine-status");
  node.className=`engine-status ${type}`.trim();
  node.querySelector("span:last-child").textContent=text;
}

function refreshButtons(){
  const confirmed=$("privacy-confirm").checked;
  $("build-button").disabled=!(engineReady&&confirmed&&$("partner-file").files[0]&&$("staff-file").files[0]&&$("master-file").files[0]&&$("build-month").value);
  $("analytics-button").disabled=!(engineReady&&confirmed&&$("base-file").files[0]&&$("analytics-month").value);
}

async function prepareEngine(){
  try{
    pyodide=await loadPyodide({indexURL:PYODIDE_URL});
    const [processor,inspector]=await Promise.all([fetch("./processor.py").then(r=>r.text()),fetch("./inspect_xlsx.py").then(r=>r.text())]);
    pyodide.FS.mkdirTree("/app");
    pyodide.FS.writeFile("/app/processor.py",processor,{encoding:"utf8"});
    pyodide.FS.writeFile("/app/inspect_xlsx.py",inspector,{encoding:"utf8"});
    await pyodide.runPythonAsync("import sys; sys.path.insert(0, '/app'); import processor");
    engineReady=true;
    setStatus("Готово. Обработка будет выполняться локально в этом браузере.","ready");
  }catch(error){
    console.error(error);
    setStatus("Не удалось подготовить обработку Excel. Обновите страницу или проверьте доступ к интернету.","error");
  }
  refreshButtons();
}

async function writeFile(file,path){
  const bytes=new Uint8Array(await file.arrayBuffer());
  pyodide.FS.writeFile(path,bytes);
}

function cleanup(paths){for(const path of paths){try{pyodide.FS.unlink(path)}catch{}}}

function download(bytes,name){
  const blob=new Blob([bytes],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),30000);
}

function showResult(id,html,error=false){const node=$(id);node.className=`result visible${error?" error":""}`;node.innerHTML=html}

async function buildBase(){
  const button=$("build-button");button.disabled=true;button.textContent="Собираем…";
  showResult("build-result","Файлы обрабатываются локально. Обычно это занимает 1–3 минуты. Не закрывайте страницу.");
  const paths=["/app/partner.xlsx","/app/staff.xlsx","/app/master.xlsx","/app/result.xlsx"];
  try{
    await Promise.all([writeFile($("partner-file").files[0],paths[0]),writeFile($("staff-file").files[0],paths[1]),writeFile($("master-file").files[0],paths[2])]);
    pyodide.globals.set("report_month",$("build-month").value);
    const raw=await pyodide.runPythonAsync("import json; json.dumps(processor.build_base('/app/partner.xlsx','/app/staff.xlsx','/app/master.xlsx','/app/result.xlsx',report_month), ensure_ascii=False)");
    const stats=JSON.parse(raw);const bytes=pyodide.FS.readFile(paths[3]);download(bytes,"Обновлённая база.xlsx");
    showResult("build-result",`База готова и скачана.<br>Сотрудники: <b>${stats.staff}</b> · Партнёр: <b>${stats.partner}</b> · Исключено удалённых: <b>${stats.deleted}</b><br>Жёлтых предложений: <b>${stats.yellow}</b> · Красных противоречий: <b>${stats.red}</b>`);
  }catch(error){console.error(error);showResult("build-result","Не удалось собрать базу. Проверьте, что выбраны правильные Excel-файлы и отчётный месяц.",true)}
  finally{cleanup(paths);button.textContent="Собрать базу";refreshButtons()}
}

async function buildAnalytics(){
  const button=$("analytics-button");button.disabled=true;button.textContent="Формируем…";
  showResult("analytics-result","Аналитика формируется локально. Обычно это занимает до одной минуты. Не закрывайте страницу.");
  const paths=["/app/base.xlsx","/app/analytics.xlsx"];
  try{
    await writeFile($("base-file").files[0],paths[0]);pyodide.globals.set("analytics_month",$("analytics-month").value);
    const raw=await pyodide.runPythonAsync("import json; json.dumps(processor.build_analytics('/app/base.xlsx','/app/analytics.xlsx',analytics_month), ensure_ascii=False)");
    const stats=JSON.parse(raw);const bytes=pyodide.FS.readFile(paths[1]);download(bytes,"Аналитика по отзывам.xlsx");
    showResult("analytics-result",`Аналитика готова и скачана.<br>За выбранный месяц: негативных — <b>${stats.negative}</b>, позитивных — <b>${stats.positive}</b>, нейтральных — <b>${stats.neutral}</b>.`);
  }catch(error){console.error(error);showResult("analytics-result","Не удалось сформировать аналитику. Проверьте выбранный файл и отчётный месяц.",true)}
  finally{cleanup(paths);button.textContent="Сформировать аналитику";refreshButtons()}
}

document.addEventListener("DOMContentLoaded",()=>{
  ["privacy-confirm","partner-file","staff-file","master-file","base-file","build-month","analytics-month"].forEach(id=>$(id).addEventListener("change",refreshButtons));
  $("build-button").addEventListener("click",buildBase);$("analytics-button").addEventListener("click",buildAnalytics);
  const now=new Date();const previous=new Date(now.getFullYear(),now.getMonth()-1,1);const value=`${previous.getFullYear()}-${String(previous.getMonth()+1).padStart(2,"0")}`;
  $("build-month").value=value;$("analytics-month").value=value;
  prepareEngine();
});
