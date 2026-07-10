# 精品货价监控数据看板 — 前端 MVP

## 项目概述

手机优先的蓝白拟态风（Neumorphism）在线数据看板，用于展示精品货价监控数据。支持 MTD/YTD 切换、模块折叠、小组+汇总表格、涨跌/达成状态颜色。

## 技术栈

- 纯 HTML + CSS + JavaScript（零依赖）
- 数据源：`data/sample_dashboard.json`（静态 JSON，未来替换为 API）
- 结构设计便于迁移至 React/Vue

## 目录结构

```
frontend/
├── index.html              # 登录页 + 看板页（单页应用）
├── styles.css              # 蓝白拟态风样式
├── app.js                  # 前端逻辑（数据加载、渲染、交互）
├── data/
│   └── sample_dashboard.json   # 模拟数据
├── README.md               # 本文件
└── design_notes.md         # 设计规范文档
```

## 运行方式

### 方式一：直接打开

用浏览器直接打开 `index.html` 即可预览。由于使用了 `fetch` 加载 JSON，建议通过本地 HTTP 服务器运行。

### 方式二：本地 HTTP 服务器（推荐）

```bash
cd frontend
python3 -m http.server 8080
# 浏览器访问 http://localhost:8080
```

或使用 Node.js：

```bash
cd frontend
npx serve .
```

### 默认密码

```
vip2026
```

密码定义在 `app.js` 中的 `PASSWORD` 变量，未来迁移至后端验证。

## 页面说明

### 1. 登录页

- 统一密码输入
- 错误提示 + 抖动反馈
- 密码正确后进入看板

### 2. 看板页

- **顶部导航**：标题、数据更新日期、退出按钮
- **MTD/YTD 切换**：月累计 / 年累计，一键切换所有数据
- **汇总概览条**：显示首个模块的全局汇总数据（实际/目标/达成率）
- **模块卡片**（可折叠）：
  - 每个模块含图标、名称、汇总行
  - 模块内按小组展示数据表格
  - 表格列：指标 | 实际 | 目标 | 达成率/同比
- **状态颜色**：
  - 🟢 绿色 = 达成 / 同比增长
  - 🔴 红色 = 未达成 / 同比下降
  - 🟡 黄色 = 接近达成（90%~100%）
  - ⚪ 灰色 = 无数据

## 移动端适配

- **手机优先设计**：默认单列布局，触控友好
- **断点策略**：
  - `< 360px`：缩小字号，紧凑布局
  - `360px ~ 768px`：标准手机/大手机
  - `768px+`：居中容器，最大宽度 720px
  - `1024px+`：桌面端，最大宽度 900px，表格行高加大
- **触摸优化**：按钮/表头点击区域 ≥ 40px，无 hover 依赖
- **粘性导航**：顶部导航栏 sticky 定位，滚动时始终可见

## 数据格式

数据源 `data/sample_dashboard.json` 结构：

```json
{
  "meta": { "title": "...", "lastUpdate": "2026-07-09", "unit": "万元" },
  "modules": [
    {
      "id": "m1",
      "name": "销售额",
      "icon": "💰",
      "summary": {
        "mtd": { "actual": 4200, "target": 4500, "yoy": 8.5 },
        "ytd": { "actual": 28500, "target": 30000, "yoy": 12.3 }
      },
      "groups": [
        {
          "name": "黄金组",
          "rows": [
            {
              "metric": "总销售额",
              "mtd": { "actual": 1800, "target": 2000, "yoy": 5.2 },
              "ytd": { "actual": 12000, "target": 13000, "yoy": 9.1 }
            }
          ]
        }
      ]
    }
  ]
}
```

## 未来迁移指南

### 迁移至 React

1. `app.js` 中的 `renderModule` / `renderGroup` 函数 → React 组件
2. `state` 对象 → `useState` / `useReducer`
3. `fetch` 逻辑 → `useEffect` + API 调用
4. CSS 直接复用，改用 CSS Modules 或 styled-components

### 迁移至 Vue

1. `index.html` 的模板部分 → `.vue` 单文件组件
2. `app.js` 的事件处理 → Vue 的事件绑定
3. `state` → Vue `ref` / `reactive`
4. CSS 直接复用

### 接入后端 API

1. 将 `DATA_URL` 改为后端 API 地址
2. 添加 token 认证（登录接口返回 JWT）
3. 添加数据刷新机制（轮询或 WebSocket）

## 已知限制

- 密码前端硬编码（MVP 阶段，后续接后端）
- 无数据缓存（每次刷新重新请求）
- 无导出功能（后续可加 Excel 导出）
- 毛利模块暂未实现（按口径要求）
