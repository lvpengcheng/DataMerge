# 登录认证 + 用户管理 + MySQL 实施方案

## 总览
为数据整合平台增加 JWT 登录认证、用户/角色/组织管理、租户数据隔离功能，使用外部 MySQL 数据库。

---

## Phase 1: 基础设施（数据库 + 模型）

### 1.1 新增依赖 (`requirements.txt`)
- SQLAlchemy 2.0、PyMySQL、cryptography
- python-jose、passlib、bcrypt

### 1.2 `.env` 新增 MySQL + JWT 配置
```
MYSQL_HOST=localhost
MYSQL_PORT=3306
MYSQL_USER=datamerge
MYSQL_PASSWORD=your_password
MYSQL_DATABASE=datamerge
DATABASE_URL=mysql+pymysql://datamerge:your_password@localhost:3306/datamerge?charset=utf8mb4

JWT_SECRET_KEY=random-secret-key
JWT_ALGORITHM=HS256
JWT_ACCESS_TOKEN_EXPIRE_MINUTES=1440
```

### 1.3 新建 `backend/database/` 模块
- `connection.py` — SQLAlchemy engine、SessionLocal、get_db 依赖
- `models.py` — 4 张表：
  | 表 | 说明 |
  |---|---|
  | users | 用户（username, password_hash, org_id, role_id, is_active...） |
  | roles | 角色（name, permissions JSON, is_system） |
  | organizations | 组织树（name, parent_id, is_active） |
  | tenant_authorizations | 租户授权（tenant_id, org_id, auth_type: owner/shared） |
- `init_db.py` — 建表 + 种子数据（admin用户、默认角色、默认组织、现有租户自动迁移）

---

## Phase 2: 认证模块 (`backend/auth/`)

- `utils.py` — 密码哈希、JWT 创建/校验
- `schemas.py` — Pydantic 请求/响应模型
- `dependencies.py` — FastAPI 依赖：
  - `get_current_user` — 从 Bearer Token 解析用户
  - `require_admin` — 管理员权限检查
  - `get_accessible_tenants` — 获取用户可访问的租户列表
- `router.py` — 认证接口：
  - `POST /api/auth/login` — 登录返回 JWT
  - `POST /api/auth/logout` — 登出
  - `GET /api/auth/me` — 当前用户信息
  - `POST /api/auth/change-password` — 修改密码

---

## Phase 3: 管理模块 (`backend/admin/router.py`)

所有接口需要 admin 权限：
- 用户管理: CRUD `/api/admin/users`，密码重置
- 角色管理: CRUD `/api/admin/roles`
- 组织管理: CRUD `/api/admin/organizations`（树形结构）
- 租户授权: `/api/admin/tenant-auth`（授权/撤销/查看）

---

## Phase 4: 前端页面

### 4.1 `frontend/static/js/auth.js` — 共享认证工具
- TOKEN/USER 的 localStorage 管理
- `AUTH.authFetch()` — 自动带 Authorization 头，401 时跳转登录
- `AUTH.requireAuth()` — 页面加载时检查登录状态
- `AUTH.renderUserInfo()` — 在 header 显示用户名 + 退出按钮

### 4.2 `login.html` + `login.js` + `login.css` — 登录页
- 沿用淡蓝色主题，居中登录卡片
- 用户名/密码输入，登录后跳转 /training

### 4.3 `admin.html` + `admin.js` + `admin.css` — 管理页
- Header 导航加"管理"标签
- 4 个 Tab：用户管理 | 角色管理 | 组织管理 | 租户授权
- 每个 Tab 有表格 + 增删改查弹窗

---

## Phase 5: 改造现有代码

### 5.1 修改 HTML 页面（training.html、compute.html）
- 引入 `auth.js`
- Header 添加"管理"导航链接 + 用户信息区域

### 5.2 修改 JS（training.js、compute.js）
- 页面加载时调用 `AUTH.requireAuth()`
- 所有 `fetch()` 替换为 `AUTH.authFetch()`

### 5.3 修改 `main.py`
- 注册 auth_router、admin_router
- 添加 /login、/admin 页面路由
- 所有现有 API 路由添加 `Depends(get_current_user)`
- 租户相关路由添加租户访问权限检查
- `/api/tenants` 和 `/api/training-history` 按组织过滤结果

---

## Phase 6: Docker 部署更新

- `docker-compose.yml` — 由于使用外部 MySQL，无需新增容器，只需确保 `.env` 里 MySQL 地址正确
- `Dockerfile` — 添加 MySQL 客户端库 (`default-libmysqlclient-dev`)

---

## 新增文件清单
```
backend/database/__init__.py
backend/database/connection.py
backend/database/models.py
backend/database/init_db.py
backend/auth/__init__.py
backend/auth/utils.py
backend/auth/schemas.py
backend/auth/dependencies.py
backend/auth/router.py
backend/admin/__init__.py
backend/admin/router.py
frontend/templates/login.html
frontend/templates/admin.html
frontend/static/js/auth.js
frontend/static/js/login.js
frontend/static/js/admin.js
frontend/static/css/login.css
frontend/static/css/admin.css
```

## 修改文件清单
```
requirements.txt
.env
Dockerfile
docker-compose.yml
backend/app/main.py
frontend/templates/training.html
frontend/templates/compute.html
frontend/static/js/training.js
frontend/static/js/compute.js
frontend/static/css/training.css
frontend/static/css/compute.css
```
