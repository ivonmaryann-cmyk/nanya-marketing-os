from __future__ import annotations

import json
import re
import uuid
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zoneinfo import ZoneInfo

from .db import db_cursor
from .product_name_catalog import CATEGORIES, SEEDS, STRUCTURES


def now():
    return datetime.now(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d %H:%M:%S")


def init_schema():
    sql = (Path(__file__).parent.parent / "migrations/automation/postgresql/0014_product_names.sql").read_text()
    with db_cursor() as conn:
        for statement in sql.split(";"):
            if statement.strip():
                conn.execute(statement)


def seed():
    """Insert missing official definitions only. Never overwrite user maintenance."""
    rows = [(cat, code, name, {}) for cat, values in SEEDS.items() for code, name in values.items()]
    source = Path(__file__).parent / "default_rules/product_names/glue.json"
    rows += [("glue", r["code"], r["name"], r["details"]) for r in json.loads(source.read_text())]
    with db_cursor() as conn:
        for cat, code, name, details in rows:
            conn.execute("""INSERT INTO product_name_mappings
                (category,code,name,details,updated_by,updated_at) VALUES (?,?,?,?,?,?)
                ON CONFLICT(category,code) DO NOTHING""",
                (cat, code, name, json.dumps(details, ensure_ascii=False), "标准初始化", now()))
    old_names={
        ('copper_pair','HHNN'): '文档示例：双面H铜重HTE（H单独重量定义待补）',
        ('board_spec','C3'): '三级', ('board_spec','T3'): '三级',
        ('board_spec','C2'): '二级', ('board_spec','T2'): '二级',
        ('roll_size','R0004970'): '自用49.7英寸',
        ('roll_size','R0004380'): '自用43.8英寸',
        ('roll_size','R000XXXX'): '外售常规49.5英寸',
    }
    for row in mappings():
        key=(row['category'],row['code'])
        if (key in old_names and row['name']==old_names[key] and
                row['revision']==1 and row['updated_by']=='标准初始化' and not row['details']):
            save_mapping(*key, SEEDS[key[0]][key[1]], {}, row['enabled'],
                         row['revision'], '业务规则更新')


def categories(product):
    return {k:v for k,v in CATEGORIES.items() if v[2] in {"shared",product}}


def mappings(category=None, keyword="", enabled_only=False):
    with db_cursor() as conn:
        return _mappings(conn, category, keyword, enabled_only)


def _mappings(conn, category=None, keyword="", enabled_only=False):
    clauses, params = [], []
    if category:
        clauses.append("category=?"); params.append(category)
    if enabled_only:
        clauses.append("enabled=1")
    rows = conn.execute("SELECT * FROM product_name_mappings" + (" WHERE " + " AND ".join(clauses) if clauses else "") + " ORDER BY category,code", params).fetchall()
    result=[]
    for row in rows:
        r=dict(row); r["details"]=json.loads(r["details"])
        if enabled_only and r['category']=='copper_type' and r['code']=='A':
            continue
        if keyword.casefold() in (r["code"]+r["name"]+json.dumps(r["details"],ensure_ascii=False)).casefold():
            result.append(r)
    return result


def save_mapping(category, code, name, details, enabled, revision, actor):
    if category not in CATEGORIES:
        raise ValueError("未知映射类别")
    code, name = code.strip().upper(), name.strip()
    width=CATEGORIES[category][1]
    if not re.fullmatch(r"[A-Z0-9]{%d}" % width, code) or not name or len(name)>200:
        raise ValueError(f"编码必须是{width}位大写字母或数字，名称必填且不超过200字")
    if category=="glass_style" and not code.isdigit():
        raise ValueError("玻布型号必须为4位数字")
    if category=="roll_size" and not re.fullmatch(r"R[0-9]{3}[A-Z0-9]{4}", code):
        raise ValueError("卷料编码须为R+3位卷长+4位幅宽编码")
    if enabled not in (0,1) or revision<0:
        raise ValueError("状态或版本无效")
    allowed=({"classification","legacy"} if category=="glue" else {"note"}) | {"aliases", "default_for"}
    details={k:str(v).strip() for k,v in details.items() if k in allowed}
    if any(len(v)>500 for v in details.values()):
        raise ValueError("备注不超过500字")
    if any(p not in {"board", "pp", ""} for p in details.get("default_for", "").split(",")):
        raise ValueError("默认适用产品无效")
    if category == "copper_type" and code == "A" and enabled:
        raise ValueError("铜箔类型A为标准预留，不可启用")
    timestamp=now()
    with db_cursor() as conn:
        row=conn.execute("SELECT * FROM product_name_mappings WHERE category=? AND code=?",(category,code)).fetchone()
        before=dict(row) if row else {}
        if revision==0:
            if before:
                raise ValueError("该编码已存在，请编辑已有记录")
            result=conn.execute("""INSERT INTO product_name_mappings
                (category,code,name,details,enabled,revision,updated_by,updated_at) VALUES (?,?,?,?,?,1,?,?)
                ON CONFLICT(category,code) DO NOTHING""",(category,code,name,json.dumps(details,ensure_ascii=False),enabled,actor,timestamp))
        else:
            result=conn.execute("""UPDATE product_name_mappings SET name=?,details=?,enabled=?,revision=revision+1,
                updated_by=?,updated_at=? WHERE category=? AND code=? AND revision=?""",
                (name,json.dumps(details,ensure_ascii=False),enabled,actor,timestamp,category,code,revision))
        if result.rowcount!=1:
            raise ValueError("记录已被其他人员修改，请刷新后重新编辑")
        after=dict(conn.execute("SELECT * FROM product_name_mappings WHERE category=? AND code=?",(category,code)).fetchone())
        conn.execute("INSERT INTO product_name_audit VALUES (?,?,?,?,?,?,?)",
                     (str(uuid.uuid4()),category,code,json.dumps(before,ensure_ascii=False),json.dumps(after,ensure_ascii=False),actor,timestamp))


def number(value, scale, width, label):
    if len(str(value))>32 or not re.fullmatch(r"[0-9]+(?:\.[0-9]+)?",str(value)):
        raise ValueError(f"{label}须填写普通十进制正数")
    try:
        n=Decimal(str(value))
    except InvalidOperation:
        raise ValueError(f"{label}须填写数字") from None
    if not n.is_finite() or n<=0 or n*scale != (n*scale).to_integral_value() or n*scale>=10**width:
        raise ValueError(f"{label}超出标准范围或精度，不能自动舍入")
    return str(int(n*scale)).zfill(width)


def build(product, values, rows):
    if product not in STRUCTURES:
        raise ValueError("未知产品")
    lookup={(r["category"],r["code"]):r for r in rows if r["enabled"] and not (r["category"]=="copper_type" and r["code"]=="A")}
    used=[]; segments=[]
    def mapped(cat, field=None):
        code=str(values.get(field or cat,"")).strip().upper()
        if cat=='copper_type' and code=='A':
            raise ValueError("铜箔类型A为预留项，不能用于转换；HTE请选择N，RTF请选择对应类型")
        record=lookup.get((cat,code))
        if not record:
            raise ValueError(f"{CATEGORIES[cat][0]}未选择、未维护或已停用，请先核对")
        used.append(record)
        return code
    glue=str(values.get("glue","")).strip().upper()
    # A legacy prefix can identify candidates, never an implicit default.
    if len(glue)!=4:
        raise ValueError("请选择完整四位胶系，旧两位码不足以确定配方与用途")
    segments.append(mapped("glue"))
    if product=="board":
        thickness=number(values.get("thickness",""),1000,4,"基板厚度(mm)")
        if int(thickness)>800 and int(thickness)%10 not in (0,5):
            raise ValueError("厚度大于0.8mm时末位须为0或5；客户特殊厚度本期不自动处理")
        segments.extend([thickness,mapped("board_spec")])
        if values.get("copper_mode")=="standard":
            segments.append(mapped("copper_pair"))
        elif values.get("copper_mode","split")=="split":
            segments.append(mapped("copper_weight","copper_top_weight")+mapped("copper_weight","copper_bottom_weight")+mapped("copper_type","copper_top_type")+mapped("copper_type","copper_bottom_type"))
        else:
            raise ValueError("请选择铜箔编码方式")
        segments.append(mapped("structure"))
    else:
        segments.extend([mapped("glass_style"),mapped("glass_type")])
        mode=values.get("pp_mode")
        encoded=number(values.get("pp_value",""),10 if mode=="rc" else 1,3,"PP厚度/RC")
        n=int(encoded)
        if mode=="thickness":
            if n>=300 or n%10 not in (0,3,5,8):
                raise ValueError("PP厚度须小于300μm且末位为0/3/5/8；其他范围待标准确认")
        elif mode=="rc":
            pass  # Explicit percentage mode, not a numeric-range inference.
        else:
            raise ValueError("请选择PP厚度或RC模式")
        segments.extend([encoded,mapped("pp_spec")])
    segments.append(mapped("marking"))
    size_mode=values.get("size_mode","sheet")
    if size_mode=="roll" and product=="pp":
        roll=str(values.get("roll_size", "")).strip().upper()
        if not re.fullmatch(r"R[0-9]{3}[A-Z0-9]{4}", roll):
            raise ValueError("卷料尺寸须为R+3位卷长+4位幅宽编码")
        # The maintained R000 record authorizes the width, not a guessed width.
        record=lookup.get(("roll_size", "R000"+roll[4:]))
        if not record:
            raise ValueError("卷料幅宽未维护或已停用，请核对")
        used.append(record)
        segments.append(roll)
    elif size_mode=="sheet":
        segments.append(number(values.get("length",""),100,4,"长度(英寸)")+number(values.get("width",""),100,4,"宽度(英寸)"))
    else:
        raise ValueError("尺寸模式无效")
    segments.append(mapped(product+"_grade"))
    if [len(s) for s in segments]!=[w for _,w in STRUCTURES[product]]:
        raise ValueError("字段长度与编码标准不一致")
    return {"code":"".join(segments),"segments":[{"label":label,"code":code} for (label,_),code in zip(STRUCTURES[product],segments)],
            "mappings":used,"standard":"new-name-v1","input":dict(values)}


def convert(product, values, actor):
    with db_cursor() as conn:
        rows=_mappings(conn,enabled_only=True)
        result=build(product,values,rows)
        if values.get("source_spec"):
            from .product_name_extract import extract
            recognized=extract(product,values["source_spec"],values.get("requirements",""),rows)
            result["recognition"]={"initial":recognized["values"],"warnings":recognized["warnings"],
                "evidence":{key:(recognized["evidence"].get(key,"自动识别") if str(value)==str(recognized["values"].get(key,"")) else "业务确认修改")
                            for key,value in values.items() if key not in {"source_spec","customer","requirements"}}}
        rid=str(uuid.uuid4())
        conn.execute("INSERT INTO product_name_results VALUES (?,?,?,?,?,?,?,?)",
            (rid,product,actor,str(values.get("source_spec",""))[:5000],json.dumps(dict(values),ensure_ascii=False),
             result["code"],json.dumps(result,ensure_ascii=False),now()))
    return rid


def history(product, actor, result_id=None):
    with db_cursor() as conn:
        rows=conn.execute("SELECT * FROM product_name_results WHERE product=? AND employee_id=?"+
            (" AND id=?" if result_id else "")+" ORDER BY created_at DESC LIMIT 100",
            (product,actor,result_id) if result_id else (product,actor)).fetchall()
    return [{**dict(r),"snapshot":json.loads(r["snapshot_json"])} for r in rows]
