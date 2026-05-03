# 设计项目进度管理系统 - 部署说明

## 项目概述
工业设计院设计项目进度管理系统，支持多用户通过浏览器访问，基于角色进行权限控制。

## 技术栈
- **后端**: Python 3.8+ / Flask
- **数据库**: SQLite（轻量级，无需额外安装）
- **前端**: HTML + CSS + JavaScript（纯前端，无框架依赖）

## 快速启动

### 方法一：使用启动脚本（推荐）
```bash
cd design-manager
chmod +x start.sh
./start.sh
```

### 方法二：手动启动
```bash
cd design-manager

# 1. 创建虚拟环境
python3 -m venv venv
source venv/bin/activate  # Linux/Mac
# venv\Scripts\activate   # Windows

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动服务
python app.py
```

启动后访问: **http://localhost:5000**

## 默认账号
| 角色 | 用户名 | 密码 | 权限 |
|------|--------|------|------|
| 管理员 | admin | admin123 | 全部权限 + 用户管理 |
| 首席工程师 | engineer | eng123 | 编辑项目进度 |
| 设计人员 | staff | staff123 | 仅查看进度 |

## 功能说明
1. **施工图计划表** - 11列，管理施工图设计进度
2. **阶段设计计划表** - 10列，跟踪各设计阶段节点
3. **技术要求计划表** - 7列，跟踪设备技术要求进度
4. **用户管理** - 管理员可增删改用户、分配角色

## 局域网部署
启动后，同一局域网内的其他电脑可通过 `http://<服务器IP>:5000` 访问。

查看本机IP：
```bash
# Linux
ip addr show | grep inet
# 或
ifconfig
```

## 云服务器部署（生产环境）

### 使用 Gunicorn（推荐）
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### 使用 Nginx 反向代理
```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    }
}
```

### 使用 systemd 服务（开机自启）
创建文件 `/etc/systemd/system/design-manager.service`：
```ini
[Unit]
Description=Design Project Manager
After=network.target

[Service]
User=www-data
WorkingDirectory=/path/to/design-manager
Environment="PATH=/path/to/design-manager/venv/bin"
ExecStart=/path/to/design-manager/venv/bin/gunicorn -w 4 -b 127.0.0.1:5000 app:app

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl enable design-manager
sudo systemctl start design-manager
```

### HTTPS 配置（使用 Let's Encrypt）
```bash
sudo apt install certbot python3-certbot-nginx
sudo certbot --nginx -d your-domain.com
```

## 安全建议
1. **修改 SECRET_KEY**: 在 `app.py` 中修改 `SECRET_KEY` 为随机字符串
2. **修改默认密码**: 首次登录后立即修改 admin 密码
3. **使用 HTTPS**: 生产环境务必配置 SSL 证书
4. **防火墙**: 仅开放必要端口（80/443）
5. **定期备份**: 备份 `design_manager.db` 数据库文件

## 数据库备份
```bash
cp design_manager.db design_manager_backup_$(date +%Y%m%d).db
```

## 数据库位置
SQLite 数据库文件自动创建在项目目录下：`instance/design_manager.db`

## 故障排查
- **端口被占用**: 修改 `app.py` 最后一行的 `port=5000` 为其他端口
- **权限不足**: 确保对项目目录有读写权限
- **依赖安装失败**: 尝试 `pip install --upgrade pip` 后重试
