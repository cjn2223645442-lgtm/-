from flask import Flask, request, redirect, url_for, render_template_string, abort, flash, send_file
import sqlite3
import uuid
import os
import sys
import tempfile
import unittest
import io
from datetime import datetime

try:
    import qrcode
except ImportError:
    qrcode = None

BASE_HTML = """
<!doctype html>
<html lang="zh-CN">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{{ title }}</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 24px; background: #f7f7f8; color: #222; }
        .wrap { max-width: 1100px; margin: 0 auto; }
        .card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); margin-bottom: 20px; }
        h1, h2, h3 { margin-top: 0; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 12px; border-bottom: 1px solid #eee; text-align: left; vertical-align: top; }
        input, button { padding: 10px 12px; border-radius: 8px; border: 1px solid #ccc; }
        button { cursor: pointer; }
        .row { display: flex; gap: 12px; flex-wrap: wrap; }
        .row > * { flex: 1; min-width: 180px; }
        .ok { color: #0a7f3f; font-weight: 700; }
        .warn { color: #c96b00; font-weight: 700; }
        .bad { color: #c90000; font-weight: 700; }
        .muted { color: #666; }
        .flash { padding: 10px 14px; border-radius: 8px; background: #eef6ff; margin-bottom: 12px; }
        .small { font-size: 13px; word-break: break-all; }
        a.btn { display: inline-block; padding: 8px 12px; text-decoration: none; border: 1px solid #ccc; border-radius: 8px; color: #222; margin-right: 8px; }
        .qr-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 16px; }
        .qr-box { border: 1px solid #eee; border-radius: 12px; padding: 16px; text-align: center; background: #fff; }
        .qr-box img { width: 180px; height: 180px; object-fit: contain; border: 1px solid #eee; padding: 8px; background: #fff; }
        @media print {
            body { background: #fff; margin: 0; }
            .card { box-shadow: none; border: 1px solid #ddd; }
            .no-print { display: none !important; }
        }
    </style>
</head>
<body>
<div class="wrap">
    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for msg in messages %}
          <div class="flash">{{ msg }}</div>
        {% endfor %}
      {% endif %}
    {% endwith %}
    {{ body|safe }}
</div>
</body>
</html>
"""


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def create_app(test_config=None):
    app = Flask(__name__)
    app.secret_key = "replace-with-your-secret-key"
    app.config.update(
        DB_PATH="seat_checkin.db",
        TESTING=False,
    )

    if test_config:
        app.config.update(test_config)

    def get_conn():
        conn = sqlite3.connect(app.config["DB_PATH"])
        conn.row_factory = sqlite3.Row
        return conn

    def init_db():
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS seats (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                seat_code TEXT UNIQUE NOT NULL,
                seat_name TEXT NOT NULL,
                qr_token TEXT UNIQUE NOT NULL,
                is_enabled INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS checkins (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                seat_id INTEGER NOT NULL,
                employee_name TEXT NOT NULL,
                employee_id TEXT,
                mobile TEXT,
                status TEXT NOT NULL,
                checkin_at TEXT NOT NULL,
                checkout_at TEXT,
                FOREIGN KEY(seat_id) REFERENCES seats(id)
            )
            """
        )
        conn.commit()
        conn.close()

    def seed_demo_data():
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) AS cnt FROM seats")
        cnt = cur.fetchone()["cnt"]
        if cnt == 0:
            demo_seats = [
                ("A-01", "A区01号工位", str(uuid.uuid4()), now_str()),
                ("A-02", "A区02号工位", str(uuid.uuid4()), now_str()),
                ("B-01", "B区01号工位", str(uuid.uuid4()), now_str()),
                ("B-02", "B区02号工位", str(uuid.uuid4()), now_str()),
            ]
            cur.executemany(
                "INSERT INTO seats (seat_code, seat_name, qr_token, created_at) VALUES (?, ?, ?, ?)",
                demo_seats,
            )
            conn.commit()
        conn.close()

    def get_seat_by_token(token):
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT * FROM seats WHERE qr_token = ? AND is_enabled = 1", (token,))
        row = cur.fetchone()
        conn.close()
        return row

    def get_active_checkin(seat_id):
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT * FROM checkins
            WHERE seat_id = ? AND status = 'occupied' AND checkout_at IS NULL
            ORDER BY id DESC LIMIT 1
            """,
            (seat_id,),
        )
        row = cur.fetchone()
        conn.close()
        return row

    def get_all_seats_with_status():
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT
                s.id,
                s.seat_code,
                s.seat_name,
                s.qr_token,
                s.is_enabled,
                c.employee_name,
                c.employee_id,
                c.mobile,
                c.checkin_at,
                c.id AS checkin_id
            FROM seats s
            LEFT JOIN checkins c
              ON c.seat_id = s.id
             AND c.status = 'occupied'
             AND c.checkout_at IS NULL
            ORDER BY s.seat_code ASC
            """
        )
        rows = cur.fetchall()
        conn.close()
        return rows

    @app.route("/")
    def home():
        rows = get_all_seats_with_status()
        body = render_template_string(
            """
            <div class="card">
                <h1>二维码签到占位系统</h1>
                <p class="muted">适用于共享工位 / 灵活工位扫码占位。支持扫码签到、查看占位状态、手动释放。</p>
                <div class="row no-print">
                    <div><a class="btn" href="{{ url_for('admin') }}">进入后台</a></div>
                    <div><a class="btn" href="{{ url_for('usage_rule') }}">查看使用规则</a></div>
                    <div><a class="btn" href="{{ url_for('history') }}">查看签到记录</a></div>
                    <div><a class="btn" href="{{ url_for('qr_center') }}">二维码中心</a></div>
                    <div><a class="btn" href="{{ url_for('qr_print') }}">批量打印页</a></div>
                </div>
            </div>

            <div class="card">
                <h2>当前工位状态</h2>
                <table>
                    <thead>
                        <tr>
                            <th>工位编码</th>
                            <th>工位名称</th>
                            <th>状态</th>
                            <th>当前占用人</th>
                            <th>签到时间</th>
                            <th>扫码入口</th>
                        </tr>
                    </thead>
                    <tbody>
                    {% for r in rows %}
                        <tr>
                            <td>{{ r.seat_code }}</td>
                            <td>{{ r.seat_name }}</td>
                            <td>
                                {% if r.employee_name %}
                                    <span class="warn">已占用</span>
                                {% else %}
                                    <span class="ok">空闲</span>
                                {% endif %}
                            </td>
                            <td>{{ r.employee_name or '-' }}</td>
                            <td>{{ r.checkin_at or '-' }}</td>
                            <td><a href="{{ url_for('scan_qr', token=r.qr_token) }}">扫码签到页</a></td>
                        </tr>
                    {% endfor %}
                    </tbody>
                </table>
            </div>
            """,
            rows=rows,
        )
        return render_template_string(BASE_HTML, title="二维码签到占位系统", body=body)

    @app.route("/rules")
    def usage_rule():
        body = """
        <div class='card'>
            <h2>建议使用规则</h2>
            <ol>
                <li>每个工位张贴唯一二维码，伙伴到达工位后扫码签到占位。</li>
                <li>如工位已被占用，页面直接提示当前占用状态，不允许重复签到。</li>
                <li>离开工位时，伙伴应主动点击“释放工位”。</li>
                <li>如需更严格管理，可增加：超时自动释放、管理员审批、钉钉单点登录。</li>
            </ol>
            <a class='btn no-print' href='/'>返回首页</a>
        </div>
        """
        return render_template_string(BASE_HTML, title="使用规则", body=body)

    @app.route("/scan/<token>", methods=["GET", "POST"])
    def scan_qr(token):
        seat = get_seat_by_token(token)
        if not seat:
            abort(404, "二维码无效或工位已停用")

        active = get_active_checkin(seat["id"])

        if request.method == "POST":
            action = request.form.get("action")

            if action == "checkin":
                if active:
                    flash("该工位当前已被占用，无法重复签到。")
                    return redirect(url_for("scan_qr", token=token))

                employee_name = request.form.get("employee_name", "").strip()
                employee_id = request.form.get("employee_id", "").strip()
                mobile = request.form.get("mobile", "").strip()

                if not employee_name:
                    flash("请输入姓名后再签到。")
                    return redirect(url_for("scan_qr", token=token))

                conn = get_conn()
                cur = conn.cursor()
                cur.execute(
                    """
                    INSERT INTO checkins (seat_id, employee_name, employee_id, mobile, status, checkin_at)
                    VALUES (?, ?, ?, ?, 'occupied', ?)
                    """,
                    (seat["id"], employee_name, employee_id, mobile, now_str()),
                )
                conn.commit()
                conn.close()
                flash("签到成功，工位已占用。")
                return redirect(url_for("scan_qr", token=token))

            if action == "checkout":
                if not active:
                    flash("当前工位为空闲状态，无需释放。")
                    return redirect(url_for("scan_qr", token=token))

                conn = get_conn()
                cur = conn.cursor()
                cur.execute(
                    """
                    UPDATE checkins
                    SET status = 'released', checkout_at = ?
                    WHERE id = ?
                    """,
                    (now_str(), active["id"]),
                )
                conn.commit()
                conn.close()
                flash("工位已成功释放。")
                return redirect(url_for("scan_qr", token=token))

            flash("无效操作。")
            return redirect(url_for("scan_qr", token=token))

        active = get_active_checkin(seat["id"])
        body = render_template_string(
            """
            <div class="card">
                <h2>{{ seat.seat_name }}（{{ seat.seat_code }}）</h2>
                {% if active %}
                    <p>当前状态：<span class="warn">已占用</span></p>
                    <p>占用人：{{ active.employee_name }}</p>
                    <p>员工编号：{{ active.employee_id or '-' }}</p>
                    <p>手机号：{{ active.mobile or '-' }}</p>
                    <p>签到时间：{{ active.checkin_at }}</p>
                    <form method="post">
                        <input type="hidden" name="action" value="checkout">
                        <button type="submit">释放工位</button>
                    </form>
                {% else %}
                    <p>当前状态：<span class="ok">空闲</span></p>
                    <form method="post">
                        <input type="hidden" name="action" value="checkin">
                        <div class="row">
                            <div><input name="employee_name" placeholder="请输入姓名" required></div>
                            <div><input name="employee_id" placeholder="请输入员工编号"></div>
                            <div><input name="mobile" placeholder="请输入手机号"></div>
                        </div>
                        <br>
                        <button type="submit">扫码签到占位</button>
                    </form>
                {% endif %}
                <br>
                <a class="btn no-print" href="/">返回首页</a>
            </div>
            """,
            seat=seat,
            active=active,
        )
        return render_template_string(BASE_HTML, title="扫码签到", body=body)

    @app.route("/qr-center")
    def qr_center():
        rows = get_all_seats_with_status()
        body = render_template_string(
            """
            <div class="card">
                <h2>二维码中心</h2>
                <p class="muted">每个工位一个二维码，扫码后直接进入对应工位的签到占位页面。</p>
                {% if not qrcode_ready %}
                    <p class="bad">当前未安装 qrcode 依赖，无法生成二维码图片。请先执行：pip install qrcode[pil]</p>
                {% endif %}
            </div>

            <div class="card">
                <table>
                    <thead>
                        <tr>
                            <th>工位编码</th>
                            <th>工位名称</th>
                            <th>签到链接</th>
                            <th>二维码</th>
                        </tr>
                    </thead>
                    <tbody>
                    {% for r in rows %}
                        <tr>
                            <td>{{ r.seat_code }}</td>
                            <td>{{ r.seat_name }}</td>
                            <td class="small">{{ url_for('scan_qr', token=r.qr_token, _external=True) }}</td>
                            <td>
                                {% if qrcode_ready %}
                                    <a class="btn" href="{{ url_for('qr_image', token=r.qr_token) }}" target="_blank">查看二维码</a>
                                {% else %}
                                    -
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                    </tbody>
                </table>
                <br>
                <a class="btn no-print" href="/">返回首页</a>
            </div>
            """,
            rows=rows,
            qrcode_ready=(qrcode is not None),
        )
        return render_template_string(BASE_HTML, title="二维码中心", body=body)

    @app.route("/qr-print")
    def qr_print():
        rows = get_all_seats_with_status()
        body = render_template_string(
            """
            <div class="card no-print">
                <h2>批量打印页</h2>
                <p class="muted">建议直接在浏览器中按 Ctrl+P 打印，贴在工位上使用。</p>
                <a class="btn" href="/">返回首页</a>
            </div>
            <div class="qr-grid">
            {% for r in rows %}
                <div class="qr-box">
                    <h3>{{ r.seat_name }}</h3>
                    <div class="muted">{{ r.seat_code }}</div>
                    {% if qrcode_ready %}
                        <img src="{{ url_for('qr_image', token=r.qr_token) }}" alt="{{ r.seat_code }}">
                    {% else %}
                        <div class="bad">未安装 qrcode 依赖</div>
                    {% endif %}
                    <div class="small" style="margin-top: 8px;">扫码签到占位</div>
                </div>
            {% endfor %}
            </div>
            """,
            rows=rows,
            qrcode_ready=(qrcode is not None),
        )
        return render_template_string(BASE_HTML, title="批量打印二维码", body=body)

    @app.route("/qr-image/<token>")
    def qr_image(token):
        seat = get_seat_by_token(token)
        if not seat:
            abort(404, "二维码无效或工位已停用")
        if qrcode is None:
            abort(500, "未安装 qrcode 依赖，请先执行 pip install qrcode[pil]")

        scan_url = url_for("scan_qr", token=token, _external=True)
        img = qrcode.make(scan_url)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return send_file(buf, mimetype="image/png")

    @app.route("/admin", methods=["GET", "POST"])
    def admin():
        if request.method == "POST":
            seat_code = request.form.get("seat_code", "").strip()
            seat_name = request.form.get("seat_name", "").strip()
            if not seat_code or not seat_name:
                flash("请填写完整的工位编码和工位名称。")
                return redirect(url_for("admin"))

            conn = get_conn()
            cur = conn.cursor()
            try:
                cur.execute(
                    "INSERT INTO seats (seat_code, seat_name, qr_token, created_at) VALUES (?, ?, ?, ?)",
                    (seat_code, seat_name, str(uuid.uuid4()), now_str()),
                )
                conn.commit()
                flash("工位新增成功。")
            except sqlite3.IntegrityError:
                flash("工位编码已存在，请勿重复新增。")
            finally:
                conn.close()
            return redirect(url_for("admin"))

        rows = get_all_seats_with_status()
        body = render_template_string(
            """
            <div class="card">
                <h2>后台管理</h2>
                <form method="post">
                    <div class="row">
                        <div><input name="seat_code" placeholder="工位编码，如 A-03" required></div>
                        <div><input name="seat_name" placeholder="工位名称，如 A区03号工位" required></div>
                        <div><button type="submit">新增工位</button></div>
                    </div>
                </form>
            </div>

            <div class="card">
                <h3>工位列表</h3>
                <table>
                    <thead>
                        <tr>
                            <th>工位编码</th>
                            <th>工位名称</th>
                            <th>占用状态</th>
                            <th>当前占用人</th>
                            <th>二维码地址</th>
                            <th>二维码图片</th>
                        </tr>
                    </thead>
                    <tbody>
                    {% for r in rows %}
                        <tr>
                            <td>{{ r.seat_code }}</td>
                            <td>{{ r.seat_name }}</td>
                            <td>
                                {% if r.employee_name %}
                                    <span class="warn">已占用</span>
                                {% else %}
                                    <span class="ok">空闲</span>
                                {% endif %}
                            </td>
                            <td>{{ r.employee_name or '-' }}</td>
                            <td class="small">{{ url_for('scan_qr', token=r.qr_token, _external=True) }}</td>
                            <td>
                                {% if qrcode_ready %}
                                    <a class="btn" href="{{ url_for('qr_image', token=r.qr_token) }}" target="_blank">打开二维码</a>
                                {% else %}
                                    -
                                {% endif %}
                            </td>
                        </tr>
                    {% endfor %}
                    </tbody>
                </table>
                <br>
                <a class="btn no-print" href="/">返回首页</a>
            </div>
            """,
            rows=rows,
            qrcode_ready=(qrcode is not None),
        )
        return render_template_string(BASE_HTML, title="后台管理", body=body)

    @app.route("/history")
    def history():
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            """
            SELECT c.*, s.seat_code, s.seat_name
            FROM checkins c
            JOIN seats s ON c.seat_id = s.id
            ORDER BY c.id DESC
            LIMIT 100
            """
        )
        rows = cur.fetchall()
        conn.close()

        body = render_template_string(
            """
            <div class="card">
                <h2>签到记录</h2>
                <table>
                    <thead>
                        <tr>
                            <th>工位</th>
                            <th>姓名</th>
                            <th>员工编号</th>
                            <th>手机号</th>
                            <th>签到时间</th>
                            <th>释放时间</th>
                            <th>状态</th>
                        </tr>
                    </thead>
                    <tbody>
                    {% for r in rows %}
                        <tr>
                            <td>{{ r.seat_name }}（{{ r.seat_code }}）</td>
                            <td>{{ r.employee_name }}</td>
                            <td>{{ r.employee_id or '-' }}</td>
                            <td>{{ r.mobile or '-' }}</td>
                            <td>{{ r.checkin_at }}</td>
                            <td>{{ r.checkout_at or '-' }}</td>
                            <td>{{ r.status }}</td>
                        </tr>
                    {% endfor %}
                    </tbody>
                </table>
                <br>
                <a class="btn no-print" href="/">返回首页</a>
            </div>
            """,
            rows=rows,
        )
        return render_template_string(BASE_HTML, title="签到记录", body=body)

    app.get_conn = get_conn
    app.init_db = init_db
    app.seed_demo_data = seed_demo_data
    app.get_seat_by_token = get_seat_by_token
    app.get_active_checkin = get_active_checkin
    return app


class SeatCheckinAppTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.temp_dir.name, "test_seat_checkin.db")
        self.app = create_app({"TESTING": True, "DB_PATH": self.db_path})
        self.app.init_db()
        self.app.seed_demo_data()
        self.client = self.app.test_client()

        conn = self.app.get_conn()
        cur = conn.cursor()
        cur.execute("SELECT qr_token FROM seats WHERE seat_code = ?", ("A-01",))
        self.token = cur.fetchone()["qr_token"]
        conn.close()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_home_page_loads(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn("二维码签到占位系统", response.get_data(as_text=True))

    def test_checkin_success(self):
        response = self.client.post(
            f"/scan/{self.token}",
            data={
                "action": "checkin",
                "employee_name": "Andy",
                "employee_id": "E001",
                "mobile": "13800000000",
            },
            follow_redirects=True,
        )
        text = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("签到成功，工位已占用。", text)
        self.assertIn("Andy", text)

    def test_duplicate_checkin_blocked(self):
        self.client.post(
            f"/scan/{self.token}",
            data={"action": "checkin", "employee_name": "Andy"},
            follow_redirects=True,
        )
        response = self.client.post(
            f"/scan/{self.token}",
            data={"action": "checkin", "employee_name": "Bob"},
            follow_redirects=True,
        )
        self.assertIn("该工位当前已被占用，无法重复签到。", response.get_data(as_text=True))

    def test_checkout_success(self):
        self.client.post(
            f"/scan/{self.token}",
            data={"action": "checkin", "employee_name": "Andy"},
            follow_redirects=True,
        )
        response = self.client.post(
            f"/scan/{self.token}",
            data={"action": "checkout"},
            follow_redirects=True,
        )
        text = response.get_data(as_text=True)
        self.assertIn("工位已成功释放。", text)
        self.assertIn("当前状态：<span class=\"ok\">空闲</span>", text)

    def test_invalid_token_returns_404(self):
        response = self.client.get("/scan/not-a-real-token")
        self.assertEqual(response.status_code, 404)

    def test_admin_add_duplicate_seat_blocked(self):
        response = self.client.post(
            "/admin",
            data={"seat_code": "A-01", "seat_name": "重复工位"},
            follow_redirects=True,
        )
        self.assertIn("工位编码已存在，请勿重复新增。", response.get_data(as_text=True))

    def test_qr_center_page_loads(self):
        response = self.client.get("/qr-center")
        self.assertEqual(response.status_code, 200)
        self.assertIn("二维码中心", response.get_data(as_text=True))

    def test_qr_print_page_loads(self):
        response = self.client.get("/qr-print")
        self.assertEqual(response.status_code, 200)
        self.assertIn("批量打印页", response.get_data(as_text=True))

    def test_qr_image_route(self):
        response = self.client.get(f"/qr-image/{self.token}")
        if qrcode is None:
            self.assertEqual(response.status_code, 500)
        else:
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.mimetype, "image/png")


def run_tests():
    suite = unittest.defaultTestLoader.loadTestsFromTestCase(SeatCheckinAppTests)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    return 0 if result.wasSuccessful() else 1


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        raise SystemExit(run_tests())

    app = create_app()
    app.init_db()
    app.seed_demo_data()

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=False,
        use_debugger=False,
        use_reloader=False,
        threaded=True,
    )


if __name__ == "__main__":
    main()
