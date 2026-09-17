import os
import tempfile
import unittest
import uuid
from datetime import date
from pathlib import Path
from unittest.mock import patch

from flask import Flask
from fangzheng_web_app import db, task_board_service as service
from fangzheng_web_app.routes import bp as main_bp
from fangzheng_web_app.task_board_routes import bp


class TaskBoardsTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        for p in [patch.dict(os.environ,{"PLATFORM_DATABASE_BACKEND":"sqlite","IDENTITY_DATABASE_BACKEND":"sqlite"}),
                  patch.object(db,"DATABASE_PATH",Path(self.temp.name)/"test.sqlite")]:
            p.start();self.addCleanup(p.stop)
        db.init_db()
        service.init_work_schema()
        db.create_user("alice",display_name="内勤甲")
        db.create_user("bob",display_name="内勤乙")
        db.create_user("leader",display_name="主管",role="admin")
        for emp in ["alice","bob","leader"]:
            db.change_user_password(emp,"test-only-password")
        self.a=self.insert_case("alice","2026-09-01T01:00:00",completed="2026-09-10T02:00:00")
        self.b=self.insert_case("bob","2026-09-10T01:00:00")
        self.c=self.insert_case("alice","2026-09-10T02:00:00",kind="order_change",status="pending_reply")
        app=Flask(__name__,template_folder=str(Path(__file__).resolve().parents[1]/"templates"))
        app.config.update(SECRET_KEY="isolated-test",TESTING=True)
        app.register_blueprint(main_bp);app.register_blueprint(bp)
        self.client=app.test_client()
        self.login("alice")

    def login(self,name):
        with self.client.session_transaction() as s:s["employee_id"]=name

    def insert_case(self,owner,created,*,completed=None,kind="new_order",status="pending_triage"):
        with db.db_cursor() as c:
            m=c.execute("INSERT INTO mail_messages(account_id,uid,subject,created_at) VALUES (?,?,?,?)",(1,str(uuid.uuid4()),"订单 <script>外部内容</script>",created)).lastrowid
            return c.execute("""INSERT INTO order_intake_cases(employee_id,mail_id,action_type,status,workflow_stage,created_at,updated_at,completed_at)
            VALUES (?,?,?,?,?,?,?,?)""",(owner,m,kind,"archived" if completed else status,"completed" if completed else "mail_triage",created,created,completed)).lastrowid

    def board(self,owner=None,**kwargs):
        return service.build_board(owner,date(2026,9,10),date(2026,9,10),**kwargs)

    def test_completion_date_and_backlog_are_independent_of_intake_date(self):
        r=self.board()
        self.assertEqual(r["summary"]["received"],2)
        self.assertEqual(r["summary"]["completed"],0)
        self.assertEqual(r['summary']['entry_completed'],1)
        self.assertEqual(r["summary"]["pending"],2)
        self.assertEqual([t['id'] for t in self.board(state='finished')['tasks']],[self.a])
        self.assertIsNone(r["summary"]["avg_minutes"])

    def test_owner_and_type_filters_apply_to_metrics_and_lists(self):
        r=self.board("alice",action_type="order_change")
        self.assertEqual(r["summary"]["pending"],1)
        self.assertEqual([t["id"] for t in r["tasks"]],[self.c])
        self.assertEqual(r["summary"]["completed"],0)
        self.assertFalse(service.load_cases("alice",self.b))

    def test_aps_ack_is_not_business_completion(self):
        with db.db_cursor() as c:
            c.execute("INSERT INTO order_interface_call_logs(case_id,employee_id,interface_key,status,created_at) VALUES (?,?,?,?,?)",(self.c,"alice","aps_order_demand_import","success","2026-09-10T03:00:00"))
        task=service.load_cases("alice",self.c)[0]
        self.assertEqual(task["stage"],"已提交 APS")
        self.assertFalse(task["is_closed"])

    def test_latest_extraction_progress_is_visible(self):
        with db.db_cursor() as c:
            for status in ["error", "running"]:
                c.execute("INSERT INTO order_entry_template_tasks(case_id,employee_id,status,started_at) VALUES (?,?,?,?)",(self.b,"bob",status,"2026-09-10T03:00:00"))
        self.assertEqual(service.load_cases("bob",self.b)[0]["stage"],"提取中")

    def test_schema_command_is_repeatable(self):
        runner=self.client.application.test_cli_runner()
        for _ in range(2):
            result=runner.invoke(args=["task_boards","init-db"])
            self.assertEqual(result.exit_code,0,result.output)

    def test_work_is_idempotent_and_unknown_data_not_zero(self):
        kwargs=dict(token=str(uuid.uuid4()),work_date="2026-09-10",minutes=12,kind="operation",note="核对")
        service.record_work(self.c,"alice",**kwargs)
        service.record_work(self.c,"alice",**kwargs)
        r=self.board("alice")
        self.assertIsNone(r["summary"]["avg_minutes"])
        self.assertEqual(len(service.load_samples('alice')[0]),1)
        with self.assertRaises(PermissionError):service.record_work(self.b,"alice",**kwargs)
        with self.assertRaises(ValueError):service.record_work(self.c,"alice",**{**kwargs,"minutes":0})

    def test_team_access_and_detail_cannot_be_bypassed(self):
        self.assertEqual(self.client.get('/task-boards/team').status_code,403)
        self.assertEqual(self.client.get(f'/task-boards/task/{self.b}').status_code,404)
        self.login("leader")
        self.assertEqual(self.client.get('/task-boards/team?start=2026-09-10&end=2026-09-10').status_code,200)
        r=self.client.get(f'/task-boards/task/{self.b}')
        self.assertEqual(r.status_code,200)
        self.assertNotIn('进入业务处理'.encode(),r.data)

    def test_views_render_escape_mail_and_reject_invalid_filters(self):
        for tab in ('tracking','efficiency'):
            r=self.client.get(f'/task-boards/mine?start=2026-09-10&end=2026-09-10&tab={tab}')
            self.assertEqual(r.status_code,200)
            self.assertNotIn(b'<script>\xe5\xa4\x96',r.data)
        self.assertEqual(self.client.get('/task-boards/mine?start=wrong').status_code,400)
        self.assertEqual(self.client.get('/task-boards/mine?start=2026-01-01&end=2026-12-01').status_code,400)
        r=self.client.get(f'/task-boards/task/{self.c}')
        self.assertEqual(r.status_code,200)
        self.assertIn(b'&lt;script&gt;',r.data)
        self.assertEqual(self.client.post(f'/task-boards/task/{self.c}',data={}).status_code,405)

    def test_missing_sampling_schema_keeps_dashboard_available(self):
        with db.db_cursor() as c:c.execute('DROP TABLE order_task_work_samples')
        self.assertEqual(self.board()['summary']['received'],2)
        self.assertEqual(self.client.get('/task-boards/mine').status_code,200)

    def test_efficiency_only_in_team_and_includes_people_without_tasks(self):
        r=self.client.get('/task-boards/mine?tab=efficiency')
        self.assertNotIn('人效统计'.encode(),r.data)
        self.assertIn('任务追踪'.encode(),r.data)
        service.record_work(self.c,'alice',token=str(uuid.uuid4()),work_date='2026-09-10',minutes=30,kind='operation',note='')
        people={p['id']:p for p in self.board()['people']}
        self.assertIsNone(people['alice']['avg_minutes'])
        self.assertEqual(people['leader']['received'],0)
        self.assertNotIn('hours',people['leader'])
        self.login('leader')
        r=self.client.get('/task-boards/team?tab=efficiency&start=2026-09-10&end=2026-09-10')
        self.assertEqual(r.status_code,200)
        self.assertIn('全员人效统计'.encode(),r.data)
        self.assertIn('内勤乙'.encode(),r.data)

    def test_utc_dates_are_compared_in_business_timezone(self):
        self.insert_case('alice','2026-09-09T17:00:00')
        self.assertEqual(self.board('alice')["summary"]["received"],2)
        self.assertEqual(service.display_time('2026-09-09T17:00:00'),'2026-09-10 01:00:00')

    def test_start_end_duration_and_completed_lines_snapshot(self):
        with db.db_cursor() as c:
            c.execute("INSERT INTO order_intake_case_events(case_id,employee_id,action,created_at) VALUES (?,?,'processing_started',?)",(self.a,'alice','2026-09-10T01:40:00'))
            c.execute("INSERT INTO order_entry_detail_events(case_id,employee_id,event_type,title,detail_json,created_at) VALUES (?,?,'domestic_order_entry_real','成功',?,?)",(self.a,'alice','{"line_count":20}','2026-09-10T02:00:00'))
        r=self.board(employee='alice',action_type='new_order')
        self.assertEqual(r['summary']['avg_minutes'],20)
        self.assertEqual(r['summary']['completed_lines'],20)
        self.assertEqual(r['summary']['timed_count'],1)
        self.assertEqual(r['people'][0]['avg_minutes'],20)
        self.assertEqual(self.board(employee='bob')['summary']['completed_lines'],0)
        self.assertEqual(self.board(action_type='order_change')['summary']['completed_lines'],0)

    def test_running_state_and_date_filter_apply_to_all_sections(self):
        with db.db_cursor() as c:
            c.execute("INSERT INTO order_intake_case_events(case_id,employee_id,action,created_at) VALUES (?,?,'processing_started',?)",(self.c,'alice','2026-09-10T03:00:00'))
        r=self.board(employee='alice',action_type='order_change',state='running')
        self.assertEqual(r['summary']['received'],1)
        self.assertEqual(r['summary']['running'],1)
        self.assertEqual(r['people'][0]['running'],1)
        self.assertEqual([t['id'] for t in r['tasks']],[self.c])
        empty=service.build_board(None,date(2026,8,1),date(2026,8,1))
        self.assertEqual(empty['summary']['received'],0)
        self.assertEqual(empty['summary']['pending'],0)
        self.assertEqual(empty['summary']['entry_completed'],0)
        self.assertEqual(empty['tasks'],[])

    def test_shared_mail_deduplicates_and_other_mail_is_counted(self):
        duplicate=self.insert_case('bob','2026-09-10T02:00:00',kind='order_change')
        with db.db_cursor() as c:
            c.execute("UPDATE mail_messages SET message_id='same-message' WHERE id IN (SELECT mail_id FROM order_intake_cases WHERE id IN (?,?))",(self.c,duplicate))
            c.execute("INSERT INTO mail_accounts(id,email,owner_employee_id,created_at,updated_at) VALUES (1,'shared@example.com','alice','2026-09-01','2026-09-01')")
            c.execute("INSERT INTO mail_messages(account_id,uid,subject,created_at) VALUES (1,'other','普通通知','2026-09-10T02:00:00')")
        r=self.board()
        self.assertEqual(r['summary']['received'],3)
        self.assertEqual(r['total'],2)
        self.assertEqual(sum(p['received'] for p in r['people']),3)
        self.assertEqual(self.board(action_type='order_change')['summary']['received'],1)

    def test_queue_records_first_start_only_and_view_does_not_start(self):
        from fangzheng_web_app.order_entry_service import queue_template_extraction
        self.client.get(f'/task-boards/task/{self.b}')
        self.assertIsNone(service.load_cases('bob',self.b)[0]['started'])
        with patch('fangzheng_web_app.order_entry_service.subprocess.Popen'):
            queue_template_extraction(self.b,'bob')
            first=service.load_cases('bob',self.b)[0]['started']
            with db.db_cursor() as c:
                c.execute("UPDATE order_entry_template_tasks SET status='error' WHERE case_id=?",(self.b,))
            queue_template_extraction(self.b,'bob')
        self.assertIsNotNone(first)
        self.assertEqual(service.load_cases('bob',self.b)[0]['started'],first)


if __name__=='__main__':unittest.main()
