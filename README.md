# PPT 智能工具台（金蝶）

## 功能

- 上传 1 个模板 PPT（保留模板已有页面）
- 上传多个内容 PPT（按上传顺序处理）
- 将内容 PPT 逐页合并到模板末尾
- 尽量保留源页样式（文本/图片/基础元素）
- 在模板和源文件尺寸不同的场景下自动做布局缩放
- 导出新的合并结果 PPT

## 本地启动

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/uvicorn app.main:app --reload --port 8000
```

浏览器打开：`http://127.0.0.1:8000`

## 测试

```bash
.venv/bin/pytest -q
```

## 接口

- `GET /`：上传页面
- `POST /merge`：
  - `template`: 单个 `.pptx`
  - `sources`: 多个 `.pptx`
  - 返回：`application/vnd.openxmlformats-officedocument.presentationml.presentation`
- `GET /ppt-import`：PPT自动入库页面
- `POST /ppt-import`：
  - `file`: 单个 `.pptx`
  - 返回：`{ chapters: [{ title, content, slide_count, ppt_base64 }] }`，按章节拆分后各章内容及 PPT
- `GET /search-fill`：PPT搜索填入页面
- `POST /search-fill`：
  - `industry`: 行业（表单字段）
  - `customer`: 客户（表单字段）
  - `model`: DeepSeek 模型（可选，默认 `deepseek-chat`）
  - `template`: 模板 `.pptx`
  - 返回：在模板末尾新增2页后的 `.pptx`
- `GET /settings`：API 可视化配置页面
- `GET /api/settings`：读取当前 DeepSeek 配置（Key 脱敏）
- `POST /api/settings`：保存 DeepSeek 配置到本地 `.env`
- `POST /api/settings/test`：测试当前配置与 DeepSeek 连通性
- `GET /api/deepseek-logs`：查看最近请求日志摘要（不含 Key）

## DeepSeek API 配置

第二模块“PPT搜索填入”默认使用 DeepSeek 接口返回结构化内容。启动前请配置环境变量：

```bash
cp .env.example .env
# 编辑 .env，填入你的 key
```

必须项：

- `DEEPSEEK_API_KEY`

可选项：

- `DEEPSEEK_BASE_URL`（默认 `https://api.deepseek.com`）
- `DEEPSEEK_MODEL`（默认 `deepseek-chat`）
- `DEEPSEEK_TIMEOUT_SECONDS`（默认 `90`）

## 部署到 GitHub

1. 在 GitHub 创建新仓库（如 `ppt-agent`）
2. 本地执行：

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<你的用户名>/<仓库名>.git
git push -u origin main
```

> 本地 `.pptx` 文件、`.env`、`.venv` 已通过 `.gitignore` 排除，不会上传。

## GitHub Actions 部署

- 推送 `main` 分支时自动运行测试
- 若配置了 `RENDER_DEPLOY_HOOK_URL` 仓库密钥，会触发 Render 重新部署
- 也可在 [Render](https://render.com) 连接本仓库，使用 `render.yaml` 一键部署（需在控制台配置 `DEEPSEEK_API_KEY`）

## 说明与边界

- 本实现优先“保留源内容样式”，采用低层 XML 复制和关系重映射。
- 对非常复杂对象（少数第三方插件对象、特殊动画链路）可能存在兼容性差异，建议在目标模板上抽样复检。
