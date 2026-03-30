# 舱单数据处理系统 - 部署指南

本文档提供了多种平台部署舱单数据处理系统的详细说明。

## 目录

- [推荐方案](#推荐方案)
- [方案一：Render 部署（推荐）](#方案一render-部署推荐)
- [方案二：Railway 部署](#方案二railway-部署)
- [方案三：Docker 部署](#方案三docker-部署)
- [方案四：Vercel 部署](#方案四vercel-部署)
- [本地开发](#本地开发)

---

## 推荐方案

| 平台 | 推荐度 | 免费额度 | 说明 |
|------|--------|----------|------|
| **Render** | ⭐⭐⭐⭐⭐ | 750小时/月 | 最适合 Python Flask 应用，支持文件上传 |
| **Railway** | ⭐⭐⭐⭐⭐ | $5/月额度 | 简单易用，一键部署 |
| **Docker** | ⭐⭐⭐⭐ | - | 适合自建服务器 |
| **Vercel** | ⭐⭐⭐ | 100GB/月 | 主要是前端平台，Python 支持有限 |

---

## 方案一：Render 部署（推荐）

Render 是最适合部署 Python Flask 应用的平台，提供免费的部署额度。

### 步骤 1: 准备代码

```bash
git init
git add .
git commit -m "Initial commit"
```

### 步骤 2: 推送到 GitHub

1. 在 GitHub 创建新仓库
2. 推送代码：

```bash
git remote add origin https://github.com/你的用户名/仓库名.git
git branch -M main
git push -u origin main
```

### 步骤 3: 在 Render 部署

1. 访问 [render.com](https://render.com)
2. 注册/登录账号
3. 点击 **"New +"** → **"Web Service"**
4. 连接你的 GitHub 仓库
5. 配置如下：

| 配置项 | 值 |
|--------|-----|
| Name | manifest-processor |
| Runtime | Python 3 |
| Build Command | `pip install -r requirements.txt` |
| Start Command | `gunicorn app:app` |
| Instance Type | **Free** |

6. 点击 **"Create Web Service"**

### 步骤 4: 等待部署

部署完成后，你会获得一个 URL，如：`https://manifest-processor.onrender.com`

---

## 方案二：Railway 部署

Railway 提供极其简单的部署体验。

### 步骤 1: 安装 Railway CLI

```bash
npm install -g @railway/cli
```

或直接使用 GitHub 登录 [railway.app](https://railway.app)

### 步骤 2: 部署

**方式 A: 通过 Web 界面**

1. 访问 [railway.app](https://railway.app)
2. 点击 **"New Project"** → **"Deploy from GitHub repo"**
3. 选择你的仓库
4. Railway 会自动检测 Python 项目并配置

**方式 B: 通过 CLI**

```bash
# 登录
railway login

# 初始化项目
railway init

# 部署
railway up
```

### 步骤 3: 获取访问地址

Railway 会自动分配一个域名，如：`https://your-app.railway.app`

---

## 方案三：Docker 部署

适合有自己服务器或 VPS 的用户。

### 步骤 1: 构建 Docker 镜像

```bash
docker build -t manifest-processor .
```

### 步骤 2: 运行容器

```bash
docker run -d -p 5000:5000 --name manifest-app manifest-processor
```

### 步骤 3: 使用 Docker Compose（推荐）

```bash
docker-compose up -d
```

### 步骤 4: 配置反向代理（可选）

使用 Nginx 反向代理：

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://localhost:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

---

## 方案四：Vercel 部署

⚠️ **注意**：Vercel 主要用于前端应用，对 Python 后端的支持有限（执行时间限制 10-60 秒）。

### 步骤 1: 安装 Vercel CLI

```bash
npm install -g vercel
```

### 步骤 2: 部署

```bash
vercel
```

按提示完成配置：

```bash
? Set up and deploy "~/your-project"? [Y/n] y
? Which scope do you want to deploy to? Your Name
? Link to existing project? [y/N] n
? What's your project's name? manifest-processor
? In which directory is your code located? ./
? Want to override the settings? [y/N] n
```

### 步骤 3: 环境变量配置

Vercel 会自动从 `requirements.txt` 安装依赖。

---

## 本地开发

### 方式 A: 直接运行

```bash
# 安装依赖
pip install -r requirements.txt

# 启动应用
python app.py
```

访问：http://localhost:5000

### 方式 B: 使用 Docker

```bash
docker-compose up
```

### 方式 C: 使用启动脚本（Windows）

```bash
start.bat
```

---

## 常见问题

### Q: 部署后无法上传文件？
**A**: 检查平台是否支持大文件上传，免费服务通常有 100MB-500MB 限制。

### Q: 中文显示乱码？
**A**: 确保系统安装了中文字体，Linux 上需要：
```bash
apt-get install fonts-wqy-microhei
```

### Q: 处理超时？
**A**: 增加超时时间配置：
```python
# app.py
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
```

### Q: 如何自定义域名？
**A**: 在各平台设置中添加自定义域名，并配置 DNS A 记录或 CNAME 记录。

---

## 文件上传限制参考

| 平台 | 免费版限制 | 付费版 |
|------|-----------|--------|
| Render | 100MB | 可协商 |
| Railway | 无硬限制 | 无限制 |
| Vercel | 4.5MB (请求体) | 50MB |
| Docker | 自定义 | 自定义 |

---

## 安全建议

1. **API 限流**: 添加速率限制防止滥用
2. **文件验证**: 验证上传文件类型
3. **HTTPS**: 生产环境强制使用 HTTPS
4. **环境变量**: 敏感信息使用环境变量

---

## 监控和日志

### Render
- 自动提供日志和监控
- 访问 Dashboard 查看

### Railway
- 内置日志查看器
- 实时性能监控

### Docker
```bash
# 查看日志
docker logs -f manifest-app

# 监控资源
docker stats manifest-app
```

---

## 更新部署

### Render/Railway
```bash
git add .
git commit -m "Update"
git push
```
平台会自动检测并重新部署。

### Docker
```bash
docker-compose pull
docker-compose up -d --build
```

---

## 需要帮助？

- [Render 文档](https://render.com/docs)
- [Railway 文档](https://docs.railway.app)
- [Docker 文档](https://docs.docker.com)
