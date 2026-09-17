from __future__ import annotations

import hmac
import secrets

from flask import Blueprint, abort, flash, redirect, render_template, request, session, url_for, jsonify
from .product_name_extract import extract, DEFAULTS

from . import product_name_service as service
from .db import is_admin_user
from .routes import require_login, current_employee
from .product_name_catalog import CATEGORIES, STRUCTURES, RULES

bp=Blueprint("product_names",__name__,url_prefix="/product-names")


@bp.post("/<product>/preview")
def preview(product):
    if product not in STRUCTURES:
        abort(404)
    payload=request.get_json(silent=True) or {}
    if not isinstance(payload,dict) or not hmac.compare_digest(session.get("product_names_csrf", "!"), str(payload.get("csrf",""))):
        abort(400)
    values=payload.get("values",{})
    if not isinstance(values,dict) or any(not isinstance(v,str) for v in values.values()) or len(str(values))>12000:
        abort(400)
    rows=service.mappings(enabled_only=True)
    response={"values":values,"evidence":{},"warnings":[]}
    if payload.get("extract"):
        if not values.get("source_spec", "").strip():
            return jsonify(error="请先输入客户规格"),422
        response=extract(product,values.get("source_spec",""),values.get("requirements",""),rows)
        response["values"].update({k:values.get(k,"") for k in ("source_spec","customer","requirements")})
    try:
        response["result"]=service.build(product,response["values"],rows)
    except ValueError as exc:
        response["error"]=str(exc)
    return jsonify(response)


@bp.before_request
def authenticate():
    return require_login()


@bp.app_context_processor
def navigation():
    return {"product_names_available":True}


@bp.cli.command("init-db")
def initialize():
    """Create additive tables and insert missing standard mappings; never replace edits."""
    service.init_schema()
    service.seed()
    print("新品名数据库与基础映射已初始化；已有映射未覆盖。")


@bp.route("/<product>",methods=["GET","POST"])
def overview(product):
    if product not in STRUCTURES:
        abort(404)
    actor=current_employee()
    admin=is_admin_user(actor)
    csrf=session.setdefault("product_names_csrf",secrets.token_urlsafe(32))
    tab=request.args.get("tab","convert")
    if tab not in {"convert","mappings","recognition","rules","history"}:
        abort(404)
    cats=service.categories(product)
    category=request.args.get("category","glue")
    if category not in cats:
        abort(400)
    error=None
    if request.method=="POST":
        if not hmac.compare_digest(csrf,request.form.get("csrf","")):
            abort(400,description="页面校验失效，请刷新后重试")
        try:
            if request.form.get("action")=="save":
                if not admin:
                    abort(403)
                service.save_mapping(category,request.form.get("code",""),request.form.get("name",""),
                    {k:request.form.get(k,"") for k in ("classification","legacy","note","aliases","default_for")},
                    int(request.form.get("enabled","1")),int(request.form.get("revision","0")),actor)
                flash("保存成功，后续转换立即使用新映射；历史结果不变。","success")
                return redirect(url_for("product_names.overview",product=product,tab=tab if tab in {"mappings","recognition"} else "mappings",category=category),code=303)
            if request.form.get("action")=="convert":
                values={k:v for k,v in request.form.items() if k not in {"csrf","action"}}
                if len(str(values))>12000:
                    raise ValueError("输入过长")
                rid=service.convert(product,values,actor)
                return redirect(url_for("product_names.overview",product=product,result=rid),code=303)
            abort(400)
        except ValueError as exc:
            error=str(exc)
    rows=service.mappings(category,request.args.get("q","")) if tab in {"mappings","recognition"} else []
    selected=None
    if tab in {"mappings","recognition"} and admin:
        if request.method=="POST" and error:
            selected={**request.form.to_dict(),"details":{k:request.form.get(k,"") for k in ("classification","legacy","note","aliases","default_for")}}
        elif request.args.get("edit"):
            selected=next((r for r in service.mappings(category) if r["code"]==request.args["edit"]),None)
            if selected is None:
                abort(404)
        elif request.args.get("new")=="1":
            selected={"code":"","name":"","revision":0,"enabled":1,"details":{}}
    records=service.history(product,actor) if tab=="history" else []
    result=None
    if request.args.get("result"):
        found=service.history(product,actor,request.args["result"])
        if not found:
            abort(404)
        result=found[0]
    choices={cat:service.mappings(cat,enabled_only=True) for cat in cats} if tab=="convert" else {}
    return render_template("product_names/overview.html",product=product,product_label="基板" if product=="board" else "PP",
        tab=tab,cats=cats,category=category,rows=rows,selected=selected,admin=admin,csrf=csrf,error=error,
        records=records,result=result,choices=choices,structures=STRUCTURES[product],rules=RULES,
        defaults=DEFAULTS, values=request.form if request.method=="POST" else {},keyword=request.args.get("q","")), (422 if error else 200)
