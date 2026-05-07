# 🚀 共享工位预约管理 Demo

一个基于 Flask + SQLite 的共享工位预约管理系统 Demo。

适用于：

- 🏢 共享办公
- 💺 灵动工位
- 🔥 热座办公
- 👥 联合办公空间
- 🧪 企业内部 Demo 展示

系统支持：

✅ 工位预约  
✅ 扫码签到  
✅ 智能找人  
✅ 工位状态查看  
✅ 后台管理  
✅ Excel 用户导入  
✅ 工位二维码  
✅ 统计报表  

---

# ✨ 项目亮点

## 🔍 智能找人

直接搜索伙伴姓名：

```text
李四
```

即可快速查看：

```text
李四当前坐在哪个工位
```

<img width="1896" height="1395" alt="image" src="https://github.com/user-attachments/assets/e54381bd-98f2-41e5-ba66-7038bd9fb5db" />

<img width="2025" height="1173" alt="image" src="https://github.com/user-attachments/assets/035d5b32-3ad7-4d55-b5d8-3c78590a8ba3" />



仅返回：

✅ 当前正在使用的工位  
❌ 不显示未来预约  

---

## 💺 工位预约

支持：

- ⏰ 自定义开始时间
- ⏰ 自定义结束时间
- ⚡ 快捷预约：

| 快捷方式 | 说明 |
|---|---|
| 30分钟 | 临时办公 |
| 1小时 | 短时会议 |
| 2小时 | 默认推荐 |
| 下班前 | 默认18:00 |
| 24点前 | 全天使用 |

---

## 📅 可视化时间轴

工位详情页支持：

- 📊 未来7天预约展示
- 🕒 时间轴可视化
- 🖱 鼠标悬停查看完整预约信息

让工位占用情况一目了然。

<img width="1236" height="1368" alt="image" src="https://github.com/user-attachments/assets/1f01a113-f875-4116-8525-b2c466961361" />

---

# 🏠 首页功能

首页支持：

✅ 工位搜索  
✅ 智能找人  
✅ 当前占用查看  
✅ 后续预约提醒  
✅ 工位状态卡片展示  

工位状态：

- 🟢 空闲
- 🟠 已占用
- 🔵 后续有预约

<img width="1353" height="1281" alt="image" src="https://github.com/user-attachments/assets/0b69322d-7a28-4e36-9e34-764273804f17" />

---

# 👤 员工侧功能

## 📌 我的记录

支持：

- 查看历史预约
- 查看当前预约
- 查看未来预约

并支持：

✅ 取消未开始预约  
✅ 释放当前工位  

<img width="1305" height="459" alt="image" src="https://github.com/user-attachments/assets/ee437f36-a153-4fb2-860b-ae045c1af91f" />

---

## 📱 扫码预约

每个工位支持：

- 独立二维码
- 扫码进入工位页面

适用于：

🏢 现场贴码使用场景

---

# 🛠 管理后台

后台采用：

✅ 左侧菜单  
✅ 右侧管理页面  

更加接近正式后台系统结构。

<img width="2538" height="1131" alt="image" src="https://github.com/user-attachments/assets/8a9d1888-9af2-4f4b-84e9-2ea2fb8dc6f4" />

---

# 👥 用户管理

支持：

- ➕ 新增用户
- ✏ 编辑用户
- ❌ 删除用户
- 🚫 禁用用户
- 🔑 重置密码
- 📥 Excel导入
- 📄 模板下载

支持筛选：

- 姓名
- 部门
- 角色
- 状态

<img width="2556" height="1331" alt="image" src="https://github.com/user-attachments/assets/32b1d66e-f39e-49f0-829b-8b4ec83c3f6f" />

---

# 💺 工位管理

支持：

✅ 创建工位  
✅ 编辑工位  
✅ 删除工位  

<img width="2556" height="1023" alt="image" src="https://github.com/user-attachments/assets/b724d3cc-72d5-4bc3-a7cd-029090186533" />

---

# 📋 预约管理

管理员可：

- 查看全部预约
- 强制取消预约
- 强制释放工位

支持筛选：

- 📅 日期
- 👤 姓名
- 🏢 部门

<img width="2556" height="627" alt="image" src="https://github.com/user-attachments/assets/6a0d74d8-7672-4912-aabf-332f620561a4" />

---

# 📊 统计报表

支持：

## 🏢 按部门统计

- 使用次数
- 使用时长
- 工位使用率

---

## 👤 按人员统计

- 预约次数
- 使用时长

---

## 💺 按工位统计

- 热门工位
- 工位使用率

支持：

✅ 日期筛选  
✅ 部门筛选  
✅ 姓名筛选  

<img width="2526" height="1365" alt="image" src="https://github.com/user-attachments/assets/f87e8027-157b-4b0b-b21f-ca6ff543151f" />

---

# 🔐 系统规则

系统内置：

- 🔒 仅允许本人释放自己的工位
- 🚫 禁止预约冲突
- ⏰ 结束时间不可超过当天24点
- 👀 未开始预约不会显示为“已占用”
- 🔍 找人仅返回当前占用情况

---

# ⚙ 技术栈

## 后端

- Python
- Flask
- SQLite

## 前端

- HTML
- CSS
- JavaScript

## 其他

- OpenPyXL（Excel导入）
- qrcode（二维码生成）

---

# 📂 项目结构

```text
.
├── seat_reservation_demo.py
├── seat_reservation.db
├── requirements.txt
└── README.md
```

---

# 🚀 安装运行

## 1️⃣ 克隆项目

```bash
git clone https://github.com/yourname/seat-reservation-demo.git

cd seat-reservation-demo
```

---

## 2️⃣ 安装依赖

```bash
pip install flask
pip install qrcode
pip install pillow
pip install openpyxl
```

或者：

```bash
pip install -r requirements.txt
```

---

## 3️⃣ 启动项目

```bash
py seat_reservation_demo.py
```

访问：

```text
http://127.0.0.1:5000
```

---

# 🔑 默认账号

## 👑 管理员

```text
账号：admin
密码：admin123
```

---

## 👤 测试员工

```text
1001 / 123456
1002 / 123456
1003 / 123456
```

---

# 📥 Excel 导入格式

Excel 列头：

```text
工号 | 姓名 | 部门 | 密码 | 角色
```

角色：

```text
admin
employee
```

---

# 🏢 适用场景

适用于：

- 🏢 企业共享办公
- 💺 灵动工位
- 🔥 热座办公
- 👥 联合办公空间
- 🚀 创业团队

---

# 🔮 后续规划

未来可继续扩展：

- 企业微信登录
- 钉钉登录
- 工位区域管理
- 会议室预约
- 审批流程
- Docker部署
- React/Vue 前后端分离
- 数据大屏

---

# 📜 License

Andy License
