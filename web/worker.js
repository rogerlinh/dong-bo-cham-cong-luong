// Chạy lõi Python (tool_dong_bo/dong_bo_luong.py) ngay trong trình duyệt bằng Pyodide.
// File người dùng chỉ nằm trong bộ nhớ của tab này, không gửi lên máy chủ nào.
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.27.2/full/";
let loiTai = null;
try{ importScripts(PYODIDE + "pyodide.js"); }
catch(e){ loiTai = "không tải được Pyodide từ " + PYODIDE + " (mạng chặn CDN?)"; }

const GLUE = `
import sys, os, json, shutil, datetime as dt
sys.path.insert(0, "/app/tool_dong_bo")
import app_web, dong_bo_luong as core

VAO, RA = "/data/vao", "/data/ra"
os.makedirs(VAO, exist_ok=True)

def quet_file(ten):
    tt = app_web.quet(os.path.join(VAO, ten))
    tt.pop("path", None)
    return json.dumps(tt, default=str, ensure_ascii=False)

def chay_tat_ca(ten_json, tach, log):
    shutil.rmtree(RA, ignore_errors=True)
    os.makedirs(RA)
    out = os.path.join(RA, f"Tong_hop_cham_cong_luong_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx")
    paths = [os.path.join(VAO, t) for t in json.loads(ten_json)]
    file_loi = []
    ds = core.chay(paths, out, log=lambda s: log(str(s)), tach_rieng=bool(tach), file_loi=file_loi)
    files = [os.path.basename(out)] + sorted(f for f in os.listdir(RA) if f != os.path.basename(out))
    return json.dumps(app_web.tom_tat(ds, files, file_loi), default=str, ensure_ascii=False)
`;

let py;
const sanSang = (async () => {
  if(loiTai) throw new Error(loiTai);
  py = await loadPyodide({indexURL: PYODIDE});
  postMessage({type: "tien_do", msg: "Đang tải thư viện đọc Excel…"});
  await py.loadPackage(["micropip"]);
  await py.runPythonAsync('import micropip\nawait micropip.install(["openpyxl==3.1.5"])');
  const nguon = await Promise.all(["dong_bo_luong.py", "app_web.py"].map(async f => {
    const r = await fetch("../tool_dong_bo/" + f + "?v=" + Date.now());
    if(!r.ok) throw new Error("Không tải được " + f);
    return [f, await r.text()];
  }));
  py.FS.mkdirTree("/app/tool_dong_bo");
  for(const [f, txt] of nguon) py.FS.writeFile("/app/tool_dong_bo/" + f, txt);
  py.runPython(GLUE);
})();
sanSang.then(() => postMessage({type: "san_sang"}),
             e => postMessage({type: "loi_khoi_dong", msg: String(e && e.message || e)}));

onmessage = async ({data: m}) => {
  try{
    await sanSang;
    if(m.type === "them"){
      py.FS.writeFile("/data/vao/" + m.ten, new Uint8Array(m.buf));
      const tt = JSON.parse(py.globals.get("quet_file")(m.ten));
      postMessage({type: "da_quet", ten: m.ten, tt});
    } else if(m.type === "bo"){
      try{ py.FS.unlink("/data/vao/" + m.ten); }catch(_){}
    } else if(m.type === "chay"){
      const log = s => postMessage({type: "log", msg: s});
      const kq = JSON.parse(py.globals.get("chay_tat_ca")(JSON.stringify(m.ten), m.tach_rieng, log));
      const files = kq.files.map(f => ({ten: f, buf: py.FS.readFile("/data/ra/" + f).buffer}));
      postMessage({type: "xong", kq, files}, files.map(f => f.buf));
    }
  }catch(e){
    postMessage({type: "loi", ten: m.ten, msg: String(e && e.message || e).split("\n").filter(Boolean).slice(-1)[0], chi_tiet: String(e && e.message || e)});
  }
};
