# 📱 QR Code Seat Check-in System（二维码工位签到占位系统）

一个基于 Flask 开发的轻量级二维码工位签到系统，适用于共享工位 / 灵活办公场景。
支持扫码占位、释放工位、后台管理、二维码生成与打印。

---

## 🚀 功能特性

* ✅ 每个工位独立二维码（扫码直达签到页）
* ✅ 实时工位占用状态展示
* ✅ 扫码签到 / 一键释放工位
* ✅ 后台管理工位（新增/查看）
* ✅ 签到记录查询
* ✅ 自动生成二维码图片
* ✅ 批量二维码打印页面（支持直接打印贴桌）
* ✅ 基于 SQLite，无需额外数据库

---

## 🧱 技术栈

* Python 3
* Flask
* SQLite
* qrcode（二维码生成）
* HTML + CSS（内嵌模板）

---

## 📦 安装与运行

### 1. 克隆项目

```bash
[git clone https://github.com/cjn2223645442-lgtm/qr-seat-checkin.git
cd qr-seat-checkin](https://github.com/cjn2223645442-lgtm/flexdesk-qr-system.git)
```

---

### 2. 安装依赖

```bash
pip install flask qrcode[pil]
```

---

### 3. 启动项目

```bash
python qr_seat_checkin_app.py
```

启动成功后访问：

```
http://127.0.0.1:5000
```

---

## 🧪 运行测试

```bash
python qr_seat_checkin_app.py test
```

---

## 📷 核心页面说明

### 🏠 首页

* 查看所有工位状态（空闲 / 已占用）
* 快速进入扫码页面

### 🔧 后台管理 `/admin`

* 新增工位
* 查看二维码链接
* 管理工位状态

### 📱 二维码中心 `/qr-center`

* 查看所有工位二维码
* 打开 / 下载二维码图片

### 🖨 批量打印 `/qr-print`

* 一页展示所有二维码
* 支持浏览器直接打印（Ctrl + P）

### 📊 签到记录 `/history`

* 查看历史签到/释放记录

---

## 📌 使用流程

1. 管理员创建工位（后台）
2. 系统自动生成二维码
3. 打印二维码并贴在工位上
4. 员工扫码进入签到页面
5. 输入姓名 → 完成占位
6. 离开时点击“释放工位”

---

## 🌐 部署说明

### 开发环境（本地运行）

直接使用 Flask 启动即可。

---

### 生产环境（推荐）

建议使用：

* Gunicorn（Linux）
* Waitress（Windows）
* Nginx（反向代理）

示例（Linux）：

```bash
gunicorn -w 2 -b 127.0.0.1:8000 wsgi:app
```

---

## 🗂 数据存储

当前使用 SQLite：

```
seat_checkin.db
```

适用于：

* 小规模团队
* 内部使用
* 低并发场景

如需扩展，可替换为：

* MySQL
* PostgreSQL

---

## ⚠️ 注意事项

* Flask 自带服务器仅用于开发环境
* 生产环境请勿使用 `app.run()`
* 建议使用 HTTPS（扫码更安全）
* 二维码地址需使用服务器公网地址或域名

---

## 🔮 后续可扩展

* 🔐 登录系统（员工身份识别）
* ⏱ 自动释放工位（超时机制）
* 🚫 防止一人占多个工位
* 📊 数据统计报表
* 🔗 对接钉钉 / 企业微信
* 📍 多区域 / 多楼层管理

---

## 🤝 贡献

欢迎提 Issue 或 PR，一起完善这个项目。

---

## 📄 License

MIT License
