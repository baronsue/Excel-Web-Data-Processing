# UI 显示结构说明

## 📋 UI 元素分类

### 1️⃣ 通用区域（始终显示）
- **模式选择卡片** (`mode-selector-card`)
  - 双表合并按钮
  - 单表处理按钮
  - 用于切换工作模式

### 2️⃣ 双表合并模式专属区域
**显示条件**: `state.mode === 'dual'`

- **双表合并区域** (`#dualTableSection`)
  - 上传左表文件(A)
  - 上传右表文件(B)
  - 选择工作表
  - 选择合并键
  - JOIN类型配置
  - 执行合并按钮

### 3️⃣ 单表处理模式专属区域
**显示条件**: `state.mode === 'single'`

- **单表处理区域** (`#singleTableSection`)
  - 上传单个文件
  - 工作表管理
  - JOIN合并设置（选择2个工作表时）
  - 合并/导出按钮

### 4️⃣ 共用区域（有数据时显示）
**显示条件**: 有处理结果数据时

#### 数据处理操作 (`#dataOperations`)
- 数据筛选
- 数据排序
- 数据清洗
- 数据透视
- 撤销/重做操作

#### 数据统计 (`#statsSection`)
- 显示表格统计信息
- 文件信息
- 行列数量统计

#### 结果预览与导出 (`#previewSection`)
- 表格预览
- 搜索功能
- 导出CSV/XLSX/JSON

#### 历史记录 (`#historySection`)
- 显示条件：有保存的历史记录时

#### 保存的模板 (`#templatesSection`)
- 显示条件：有保存的模板时

## 🎨 主题切换

### 双表模式（紫色主题）
- Background: `#f5e6ff → #e6d9ff → #d9ccff`
- Primary: `#9370db`
- Body class: `mode-dual`

### 单表模式（蓝色主题）
- Background: `#e0f2fe → #bae6fd → #7dd3fc`  
- Primary: `#3b82f6`
- Body class: `mode-single`

## 🔧 关键函数

### `switchMode(mode)`
切换工作模式，控制UI区域的显示/隐藏

### `showSharedSections()`
显示共用区域（统计、预览）

### `hideSharedSections()`
隐藏共用区域

### `renderTable(header, rows)`
渲染表格数据，有数据时自动显示共用区域

### `renderDataStats()`
渲染统计信息，有统计数据时显示统计区域

## 📝 HTML 结构标记

```html
<!-- 模式专属区域 -->
<section id="dualTableSection">...</section>      <!-- 双表模式 -->
<section id="singleTableSection">...</section>    <!-- 单表模式 -->

<!-- 共用区域（添加 shared-section 类和 hidden 属性） -->
<section class="card shared-section" id="dataOperations" hidden>...</section>
<section class="card shared-section" id="statsSection" hidden>...</section>
<section class="card shared-section" id="previewSection" hidden>...</section>
<section class="card shared-section" id="historySection" hidden>...</section>
<section class="card shared-section" id="templatesSection" hidden>...</section>
```

## ✅ 测试要点

1. **切换到双表模式**: 应该只看到双表合并区域，单表区域隐藏
2. **切换到单表模式**: 应该只看到单表处理区域，双表区域隐藏
3. **无数据状态**: 统计和预览区域应该隐藏
4. **有数据状态**: 统计和预览区域应该显示，并有淡入动画
5. **主题切换**: 切换模式时背景、按钮颜色应该平滑过渡

