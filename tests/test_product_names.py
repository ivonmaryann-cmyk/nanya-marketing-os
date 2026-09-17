import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from flask import Flask
from fangzheng_web_app import db, product_name_service as service
from fangzheng_web_app.product_name_routes import bp
from fangzheng_web_app.routes import bp as main_bp
from fangzheng_web_app.product_name_extract import extract


class ProductNamesTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        for p in [patch.dict(os.environ,{"PLATFORM_DATABASE_BACKEND":"sqlite","IDENTITY_DATABASE_BACKEND":"sqlite"}),
                  patch.object(db,"DATABASE_PATH",Path(self.temp.name)/"test.sqlite")]:
            p.start();self.addCleanup(p.stop)
        db.init_db();service.init_schema();service.seed()
        for name,role in [("admin","admin"),("alice","user"),("bob","user")]:
            db.create_user(name,display_name=name,role=role)
            db.change_user_password(name,"test-only-password")
        app=Flask(__name__,template_folder=str(Path(__file__).resolve().parents[1]/"templates"))
        app.config.update(SECRET_KEY="isolated",TESTING=True)
        app.register_blueprint(main_bp);app.register_blueprint(bp)
        self.client=app.test_client();self.login("admin")

    def login(self,actor):
        with self.client.session_transaction() as s:s["employee_id"]=actor

    def pp(self,**kw):
        return dict(glue="2BNN",glass_style="7628",glass_type="C",pp_mode="rc",pp_value="41",
                    pp_spec="XX",marking="W",size_mode="roll",roll_size="R0004970",pp_grade="A0",**kw)

    def build(self,values,product="pp"):
        return service.build(product,values,service.mappings())

    def test_document_examples(self):
        p=self.pp()
        self.assertEqual(self.build(p)["code"],"2BNN7628C410XXWR0004970A0")
        p.update(pp_mode="thickness",pp_value="203")
        self.assertEqual(self.build(p)["code"],"2BNN7628C203XXWR0004970A0")
        p.update(pp_mode="rc",pp_value="45",pp_spec="GA",roll_size="R000XXXX")
        self.assertEqual(self.build(p)["code"],"2BNN7628C450GAWR000XXXXA0")
        p.update(pp_value="50",pp_spec="XX",size_mode="sheet",length="20.40",width="24.41")
        self.assertEqual(self.build(p)["code"],"2BNN7628C500XXW20402441A0")
        b=dict(glue="2BNN",thickness="0.103",board_spec="C3",copper_mode="standard",copper_pair="HHNN",structure="A",marking="Q",size_mode="sheet",length="37",width="49",board_grade="A0")
        self.assertEqual(self.build(b,"board")["code"],"2BNN0103C3HHNNAQ37004900A0")
        b.update(thickness="0.2",marking="W",length="43")
        self.assertEqual(self.build(b,"board")["code"],"2BNN0200C3HHNNAW43004900A0")

    def test_automatic_specification_and_live_preview(self):
        rows=service.mappings()
        board=extract("board", 'NY2150 0.103mm H/H HTE 37"x49" 印字', "", rows)
        self.assertEqual(service.build("board",board['values'],rows)['code'],"2BNN0103C3HHNNAQ37004900A0")
        pp=extract("pp",'NY2150P 7628 RC50% 20.40x24.41',"",rows)
        self.assertEqual(service.build("pp",pp['values'],rows)['code'],"2BNN7628C500XXW20402441A0")
        car=extract("pp",'NY2150P 7628 RC50% 20.40x24.41',"汽车板",rows)
        self.assertEqual(car['values']['glue'],'2BNA')
        screenshot_spec='NY2150；0.103mm；C3；Hoz HTE铜箔；第一结构；印字；37.00×49.00；A级'
        parsed=extract('board',screenshot_spec,'',rows)
        self.assertEqual(service.build('board',parsed['values'],rows)['code'],'2BNN0103C3HHNNAQ37004900A0')
        self.assertTrue(any('双面' in w for w in parsed['warnings']))
        self.assertNotIn('A',[r['code'] for r in service.mappings('copper_type',enabled_only=True)])
        invalid=dict(parsed['values'],copper_mode='split',copper_top_weight='1',copper_bottom_weight='1',copper_top_type='A',copper_bottom_type='A')
        with self.assertRaisesRegex(ValueError,'预留项'):
            service.build('board',invalid,rows)
        self.client.get('/product-names/pp')
        with self.client.session_transaction() as s: csrf=s['product_names_csrf']
        response=self.client.post('/product-names/pp/preview',json={'csrf':csrf,'extract':True,'values':{'source_spec':'NY2150P 7628 RC50% 20.40x24.41'}})
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.json['result']['code'],"2BNN7628C500XXW20402441A0")
        self.assertFalse(service.history('pp','admin'))
        values=response.json['values'];values['pp_value']='45'
        response=self.client.post('/product-names/pp/preview',json={'csrf':csrf,'values':values})
        self.assertIn('C450XX',response.json['result']['code'])
        self.assertEqual(self.client.post('/product-names/pp/preview',json={'csrf':'bad'}).status_code,400)

    def test_recognition_maintenance_and_conflicting_defaults(self):
        r=service.mappings('pp_spec','GA')[0]
        service.save_mapping('pp_spec','GA',r['name'],{'aliases':'快速凝胶','default_for':'pp'},1,r['revision'],'admin')
        result=extract('pp','NY2150P 7628 RC50% 20.40x24.41 快速凝胶','',service.mappings())
        self.assertEqual(result['values']['pp_spec'],'GA')
        result=extract('pp','NY2150P 7628 RC50% 20.40x24.41','',service.mappings())
        self.assertNotIn('pp_spec',result['values'])
        self.assertTrue(result['warnings'])

    def test_missing_ambiguous_and_boundaries(self):
        for field,value in [("glue","2B"),("glue","XXXX"),("pp_value","NaN"),("pp_value","Infinity"),("pp_value","100"),("pp_value","41.01"),("pp_mode","bad"),("size_mode","bad")]:
            p=self.pp();p[field]=value
            with self.subTest(field=field,value=value),self.assertRaises(ValueError):self.build(p)
        for v in ["300","202","0","-1"]:
            p=self.pp();p.update(pp_mode="thickness",pp_value=v)
            with self.assertRaises(ValueError):self.build(p)
        p=self.pp();p.update(size_mode="sheet",length="20.405",width="24")
        with self.assertRaises(ValueError):self.build(p)

    def test_seed_master_and_repeat_does_not_overwrite(self):
        rows=service.mappings("glue");self.assertEqual(len(rows),282)
        r=next(x for x in rows if x["code"]=="2BNN")
        service.save_mapping("glue","2BNN","修改名称",r["details"],1,r["revision"],"admin")
        service.seed()
        self.assertEqual(service.mappings("glue","修改名称")[0]["revision"],2)

    def test_business_confirmed_rules(self):
        rows=service.mappings()
        for width,suffix in [('49.7','4970'),('43.8','4380'),('49','XXXX')]:
            result=extract('pp',f'NY2150P 7628 40% 卷料 300M 幅宽{width}', '', rows)
            self.assertEqual(result['values']['roll_size'],'R300'+suffix)
            self.assertIn('C400XXWR300'+suffix,service.build('pp',result['values'],rows)['code'])
        result=extract('pp','NY2150P 7628 48% 卷料 幅宽49.7','',rows)
        self.assertEqual(result['values']['roll_size'],'R0004970')
        result=extract('board','NY2150 0.103mm 1oz/Hoz RTF3 37x49','',rows)
        self.assertIn('H1GG',service.build('board',result['values'],rows)['code'])
        result=extract('board','NY2150 0.103mm 2oz/Hoz RTF3 37x49','',rows)
        with self.assertRaises(ValueError): service.build('board',result['values'],rows)
        result=extract('pp','NY2150P 7628 400UM 卷料 幅宽49','',rows)
        with self.assertRaises(ValueError): service.build('pp',result['values'],rows)
        p=self.pp();p['roll_size']='R300ZZZZ'
        with self.assertRaises(ValueError): self.build(p)

    def test_seed_upgrades_only_untouched_records(self):
        with db.db_cursor() as conn:
            conn.execute("UPDATE product_name_mappings SET name='三级' WHERE category='board_spec' AND code='C3'")
        service.seed()
        self.assertEqual(service.mappings('board_spec','C3')[0]['name'],'芯厚 C3')
        r=service.mappings('board_spec','T3')[0]
        service.save_mapping('board_spec','T3','业务维护名称',{},1,r['revision'],'admin')
        service.seed()
        self.assertEqual(service.mappings('board_spec','T3')[0]['name'],'业务维护名称')

    def test_concurrent_save_and_disabled_mapping(self):
        r=service.mappings("glue","2BNN")[0]
        service.save_mapping("glue",r["code"],r["name"],r["details"],0,1,"admin")
        with self.assertRaisesRegex(ValueError,"其他人员"):
            service.save_mapping("glue",r["code"],"stale",{},1,1,"admin")
        with self.assertRaises(ValueError):self.build(self.pp())
        with db.db_cursor() as c:self.assertEqual(c.execute("SELECT COUNT(*) AS n FROM product_name_audit").fetchone()["n"],1)

    def test_historical_snapshot_and_owner_isolation(self):
        rid=service.convert("pp",self.pp(),"alice")
        service.save_mapping("glue","2BNN","新名称",{},1,1,"admin")
        r=service.history("pp","alice",rid)[0]
        self.assertEqual(r["snapshot"]["mappings"][0]["name"],"NY2150")
        self.assertFalse(service.history("pp","bob",rid))
        self.login("bob");self.assertEqual(self.client.get('/product-names/pp?result='+rid).status_code,404)

    def test_pages_permissions_csrf_validation_and_escape(self):
        for product in ("board","pp"):
            for tab in ("convert","mappings","recognition","rules","history"):
                r=self.client.get(f'/product-names/{product}?tab={tab}')
                self.assertEqual(r.status_code,200,r.data[:500])
        r=self.client.get('/product-names/board?tab=mappings&edit=2BNN')
        self.assertEqual(r.status_code,200)
        with self.client.session_transaction() as s:csrf=s['product_names_csrf']
        data=dict(action="save",code="2BNN",name='<script>alert(1)</script>',revision=1,enabled=1,csrf=csrf)
        self.assertEqual(self.client.post('/product-names/board?tab=mappings',data={**data,"csrf":"bad"}).status_code,400)
        self.login("alice")
        self.assertEqual(self.client.post('/product-names/board?tab=mappings',data=data).status_code,403)
        self.login("admin")
        self.assertEqual(self.client.post('/product-names/board?tab=mappings',data=data).status_code,303)
        r=self.client.get('/product-names/board?tab=mappings')
        self.assertNotIn(b'<script>alert(1)</script>',r.data)
        self.assertIn(b'&lt;script&gt;',r.data)
        r=self.client.post('/product-names/pp',data={"action":"convert","csrf":csrf,**self.pp()})
        self.assertEqual(r.status_code,303)
        self.assertEqual(self.client.get(r.location).status_code,200)


if __name__=='__main__':unittest.main()
