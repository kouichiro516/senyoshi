
# -*- coding: utf-8 -*-
import io, json, shutil, csv, datetime as dt
from pathlib import Path
from flask import Flask, render_template, request, send_file, make_response, jsonify, Response
from openpyxl import load_workbook


# ==== Safe helpers injected ====
def _resolve_mode_key(mapping:dict, mode:str)->str:
    aliases = {
        "increase": ["increase","inc","増車","増"],
        "decrease": ["decrease","dec","減車","減"]
    }
    for k in aliases.get(mode, [mode]):
        if k in mapping: return k
    # try loose compare ignoring spaces
    norm = lambda s: "".join(str(s).split())
    for k in list(mapping.keys()):
        if norm(k) in [norm(x) for x in aliases.get(mode, [mode])]:
            return k
    raise ValueError("mapping.yaml に '増車/減車'（increase/decrease）に対応する設定が見つかりません。")

def _ensure_mode_keys(mapping:dict)->dict:
    # make a shallow copy and ensure canonical English keys exist
    mp = dict(mapping)
    # promote JP keys to EN
    if "increase" not in mp:
        for k in ("増車","増"):
            if k in mp: mp["increase"] = mp[k]; break
    if "decrease" not in mp:
        for k in ("減車","減"):
            if k in mp: mp["decrease"] = mp[k]; break
    # if still missing one side, mirror the other so動作が止まらない（値は同一）
    if "increase" not in mp and "decrease" in mp:
        mp["increase"] = mp["decrease"]
    if "decrease" not in mp and "increase" in mp:
        mp["decrease"] = mp["increase"]
    return mp

def _safe_load_wb(file_or_bytes):
    try:
        if hasattr(file_or_bytes,"read"):
            return load_workbook(io.BytesIO(file_or_bytes.read()), data_only=True)
        else:
            return load_workbook(file_or_bytes, data_only=True)
    except (BadZipFile, InvalidFileException) as e:
        raise ValueError("連絡書は.xlsx形式のExcelファイルのみ対応です。ファイルを確認してください。") from e
# ==== end helpers ====
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException
import yaml, xlwings as xw

app = Flask(__name__)
ROOT = Path(__file__).resolve().parent
SRC = ROOT / "templates_src"
OUTPUT = ROOT / "outputs"
DATA = ROOT / "data"
OUTPUT.mkdir(exist_ok=True); DATA.mkdir(exist_ok=True)
TPL_NAME = "申請書共通.xlsx"

def js_alert(msg:str):
    html = f"""<script>alert({json.dumps(msg)});history.back();</script>"""
    r = make_response(html); r.headers["Content-Type"]="text/html; charset=utf-8"; return r

def norm_spaces(s:str)->str:
    if s is None: return ""
    for z in ("\u00A0","\u200B","\u200C","\u200D"): s=s.replace(z,"")
    return s.replace(" ","").replace("\t","").replace("\u3000","").strip()

def load_mapping()->dict:
    with (ROOT/"mapping.yaml").open("r",encoding="utf-8") as f: return yaml.safe_load(f)

def read_contact_values(file_or_bytes, mode:str, mapping:dict)->dict:
    src = mapping[_resolve_mode_key(mapping, mode)]
    if hasattr(file_or_bytes,"read"): wb = _safe_load_wb(file_or_bytes)
    else: wb = _safe_load_wb(file_or_bytes)
    ws = wb.worksheets[0]
    def at(a): 
        if not a: return ""
        v = ws[a].value; return str(v).strip() if v is not None else ""
    out = {"user_name":at(src.get("user_name")),"office_name":at(src.get("office_name")),
           "era":at(src.get("era")),"year":at(src.get("year")),"capacity":at(src.get("capacity"))}
    wb.close(); return out

def parse_office_display(name:str):
    name=(name or "").strip()
    if not name: return "","",""
    if "本社" in name and "営業所" in name: return "本社営業所","", "本社"
    if name.endswith("営業所"): loc=name[:-3]; return f"{loc}営業所", loc, "営業所"
    if name.endswith("支店"): loc=name[:-2]; return f"{loc}支店", loc, "支店"
    return name, name, ""

# DB I/O
def _db_path()->Path: return DATA/"counter.json"

def load_db_raw()->dict:
    p=_db_path()
    if not p.exists(): return {"headers":[], "records":[]}
    with p.open("r",encoding="utf-8") as f: raw=json.load(f)
    if isinstance(raw,dict) and "area_table" in raw and isinstance(raw["area_table"],dict):
        at=raw["area_table"]; return {"headers":at.get("headers",[]), "records":at.get("records",[])}
    return {"headers":raw.get("headers",[]), "records":raw.get("records",[])}

def save_db_raw(db_flat:dict)->None:
    p=_db_path()
    base={}
    if p.exists():
        try:
            with p.open("r",encoding="utf-8") as f: base=json.load(f) or {}
        except Exception: base={}
    base["area_table"]={"headers":db_flat.get("headers",[]),"records":db_flat.get("records",[])}
    with p.open("w",encoding="utf-8") as f: json.dump(base,f,ensure_ascii=False,indent=2)

def find_company_index(db_flat:dict, company:str, office_display:str)->int:
    company_n=norm_spaces(company)
    disp, loc, divi = parse_office_display(office_display)
    loc_n=norm_spaces(loc)
    headers=db_flat.get("headers",[]); recs=db_flat.get("records",[])
    def val(rec,key): return rec.get(key,"") if isinstance(rec,dict) else rec[headers.index(key)]
    for i,r in enumerate(recs):
        comp=norm_spaces(val(r,"会社名")); div=str(val(r,"本社/営業所/支店") or "").strip()
        locv=norm_spaces(str(val(r,"営業所所在地") or ""))
        if comp==company_n and (div==(divi or "本社")) and ((div=="本社" and locv=="") or (div!="本社" and locv==loc_n)):
            return i
    return -1

def get_type_counts_from_row(rec:dict)->dict:
    def num(v): 
        try: return int(float(str(v).strip()))
        except: return 0
    return {"普通":num(rec.get("普通",0)), "小型":num(rec.get("小型",0)),
            "けん引":num(rec.get("けん引",0)), "被けん引":num(rec.get("被けん引用", rec.get("被けん引",0)))}

def update_db_counts_inplace(db_flat:dict, idx:int, car_type:str, delta:int)->dict:
    rec=db_flat["records"][idx]; now=get_type_counts_from_row(rec)
    rec[car_type]=int(now.get(car_type,0))+int(delta)
    total=sum(get_type_counts_from_row(rec).values())
    if "合計" in rec: rec["合計"]=total
    save_db_raw(db_flat); return get_type_counts_from_row(rec)

def put_old_new_counts(xlsx:Path, office_disp:str, old_counts:dict, new_counts:dict):
    wb = _safe_load_wb(xlsx); ws=wb.worksheets[0]
    write_row=None
    for r in range(5,10):
        if not (str(ws[f"B{r}"].value or "").strip()): write_row=r; break
    if write_row is None: write_row=9
    ws[f"B{write_row}"]=office_disp
    ws[f"K{write_row}"]=int(old_counts.get("普通",0)); ws[f"N{write_row}"]=int(old_counts.get("小型",0))
    ws[f"O{write_row}"]=int(old_counts.get("けん引",0)); ws[f"Q{write_row}"]=int(old_counts.get("被けん引",0))
    ws[f"D{write_row}"]=int(new_counts.get("普通",0)); ws[f"E{write_row}"]=int(new_counts.get("小型",0))
    ws[f"G{write_row}"]=int(new_counts.get("けん引",0)); ws[f"H{write_row}"]=int(new_counts.get("被けん引",0))
    wb.save(xlsx); wb.close()

@app.get("/")
def index(): return render_template("index.html")

@app.get("/db")
def db_get(): return jsonify(load_db_raw())

@app.post("/db/save")
def db_save():
    try: payload=request.get_json(force=True)
    except Exception: return "invalid json", 400
    if not isinstance(payload,dict) or "records" not in payload: return "bad payload", 400
    headers=payload.get("headers",[]); recs=payload.get("records",[])
    if not headers:
        keys=[]; 
        for r in recs:
            if isinstance(r,dict):
                for k in r.keys():
                    if k not in keys: keys.append(k)
        headers=keys
    num_cols={"普通","小型","けん引","被けん引","合計"}
    normed=[]
    for r in recs:
        row=dict(r) if isinstance(r,dict) else {}
        for k in list(row.keys()):
            v=row[k]
            if k in num_cols:
                try: row[k]=int(float(str(v).strip())) if str(v).strip()!="" else 0
                except Exception: row[k]=0
        normed.append(row)
    save_db_raw({"headers":headers,"records":normed}); return "ok", 200

@app.get("/db/csv")
def db_csv():
    flat = load_db_raw()
    headers = flat.get("headers", [])
    recs = flat.get("records", [])

    def generate():
        yield ",".join(headers) + "\n"
        for r in recs:
            row = [str((r.get(h, "") if isinstance(r, dict) else "")) for h in headers]
            yield ",".join(row) + "\n"

    return Response(
        generate(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=counter.csv"},
    )

@app.post("/process")
def process():
    car_type=request.form.get("car_type","普通").strip()
    inc_files=[f for f in request.files.getlist("increase_files") if getattr(f,"filename","")]
    dec_files=[f for f in request.files.getlist("decrease_files") if getattr(f,"filename","")]
    if not inc_files and not dec_files: return js_alert("増車・減車どちらかのファイルをアップロードしてください。")

    # ==== per-file meta (from hidden JSON) ====
    import json as _json
    def _load_meta(s):
        try:
            v = _json.loads(s or "[]")
            return v if isinstance(v, list) else []
        except Exception:
            return []
    def _meta_dict(meta_list):
        return { (m.get("name"), int(m.get("size",0) or 0)): m for m in meta_list if isinstance(m, dict) }
    def _file_key(fs):
        return (getattr(fs, "filename", ""), int(getattr(fs, "content_length", 0) or 0))
    inc_meta = _meta_dict(_load_meta(request.form.get("inc_meta")))
    dec_meta = _meta_dict(_load_meta(request.form.get("dec_meta")))
    WEIGHT_MAP = {
        "to_2t": "2.0トンまで",
        "2t_long": "2.0トンロング",
        "over_2t_long_to_7_5t": "2.0トンロング超～ 7.5トンまで",
        "over_7_5t": "7.5トンを超えるもの",
    }

    MAP=_ensure_mode_keys(load_mapping())
    metas=[]
    for f in inc_files: metas.append(("increase", read_contact_values(f,"increase",MAP)))
    for f in dec_files: metas.append(("decrease", read_contact_values(f,"decrease",MAP)))

    companies={m[1]["user_name"] for m in metas}; offices={m[1]["office_name"] for m in metas}
    if len(companies)!=1 or len(offices)!=1: return js_alert("アップロード内で会社名/営業所名が混在しています。1社1拠点で実行してください。")
    company_name=next(iter(companies)); office_name=next(iter(offices))
    office_disp,_,_=parse_office_display(office_name)

    db_flat=load_db_raw(); idx=find_company_index(db_flat, company_name, office_name)
    if idx<0: return js_alert(f"会社・拠点がDBに見つかりません。会社={company_name} / 営業所={office_name}")

    old_counts=get_type_counts_from_row(db_flat["records"][idx])
    # per-file shortage check by car_type
    from collections import Counter as _Counter
    _dec_by_type = _Counter()
    for _f in dec_files:
        _m = dec_meta.get(_file_key(_f), {})
        _t = (_m.get("car_type") or car_type).strip()
        _dec_by_type[_t] += 1
    for _t, _n in _dec_by_type.items():
        if old_counts.get(_t, 0) - _n < 0:
            return js_alert(f"現在、{company_name} {office_name} の {_t} が不足します（減車{_n}台は不可）。", status=409)

    tpl_path=SRC/TPL_NAME
    if not tpl_path.exists(): return js_alert(f"テンプレートが見つかりません: {tpl_path}")
    work=OUTPUT/f"work_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    shutil.copyfile(tpl_path, work)

    new_counts=old_counts.copy()
    if inc_files:
        from collections import Counter as _Counter
        _inc_by_type = _Counter()
        for _f in inc_files:
            _m = inc_meta.get(_file_key(_f), {})
            _t = (_m.get("car_type") or car_type).strip()
            _inc_by_type[_t] += 1
        for _t, _n in _inc_by_type.items():
            new_counts = update_db_counts_inplace(db_flat, idx, _t, +_n)
    if dec_files:
        from collections import Counter as _Counter
        _dec_by_type = _Counter()
        for _f in dec_files:
            _m = dec_meta.get(_file_key(_f), {})
            _t = (_m.get("car_type") or car_type).strip()
            _dec_by_type[_t] += 1
        for _t, _n in _dec_by_type.items():
            new_counts = update_db_counts_inplace(db_flat, idx, _t, -_n)

    put_old_new_counts(work, office_disp, old_counts, new_counts)

    modes=(["increase"]*len(inc_files))+(["decrease"]*len(dec_files))
    # toggle_detail_circles機能は一時的に無効化
    # try: toggle_detail_circles(work, modes)
    # except Exception as e: print(f"[warn] toggle_detail_circles skipped: {e}", flush=True)

    out_name=OUTPUT/f"事業計画変更届出_{office_disp or '不明営業所'}.xlsx"
    cnt=2
    while out_name.exists():
        out_name=out_name.with_stem(out_name.stem+f"({cnt})"); cnt+=1
    shutil.move(work, out_name)
    return send_file(out_name, as_attachment=True, download_name=out_name.name)

if __name__=="__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
