# UI 模式分离验证指南

## 🎯 预期行为

### 初始状态（双表模式）
打开页面后，您应该看到：
- ✅ **模式选择卡片**（紫色激活的"双表合并"按钮）
- ✅ **双表合并区域**（包含左表、右表上传）
- ❌ **单表处理区域**（完全不可见）
- ❌ **数据统计**（不可见）
- ❌ **结果预览**（不可见）

### 切换到单表模式后
点击"单表处理"按钮后，您应该看到：
- ✅ **模式选择卡片**（蓝色激活的"单表处理"按钮）
- ❌ **双表合并区域**（完全不可见）
- ✅ **单表处理区域**（包含单表上传、工作表管理）
- ❌ **数据统计**（不可见）
- ❌ **结果预览**（不可见）

### 处理数据后
无论哪种模式，在上传并处理数据后：
- ✅ **对应模式的区域**（继续显示）
- ✅ **数据统计区域**（显示统计信息）
- ✅ **结果预览区域**（显示表格数据）
- ✅ **数据处理操作**（显示筛选、排序等功能）

## 🔍 验证步骤

### 方法1: 视觉验证
1. 打开 `index.html`
2. 观察页面，应该只看到：
   - 顶部标题和按钮
   - 模式选择卡片
   - 双表合并区域（两个上传框）
3. 点击"单表处理"按钮
4. 页面应该切换为蓝色主题
5. 双表区域消失，单表区域出现
6. 点击"双表合并"按钮
7. 页面切换回紫色主题
8. 单表区域消失，双表区域出现

### 方法2: 使用调试按钮
1. 打开 `index.html`
2. 点击右上角的 "🐛 调试" 按钮
3. 打开浏览器控制台 (F12 或 Cmd+Option+I)
4. 查看输出信息，确认：
   ```
   当前模式: dual
   双表区域 hidden: false
   单表区域 hidden: true
   统计区域 hidden: true
   预览区域 hidden: true
   ```

5. 点击"单表处理"按钮
6. 再次点击 "🐛 调试"
7. 确认输出变为：
   ```
   当前模式: single
   双表区域 hidden: true
   单表区域 hidden: false
   统计区域 hidden: true
   预览区域 hidden: true
   ```

### 方法3: 浏览器检查元素
1. 打开 `index.html`
2. 右键点击页面 → 检查元素
3. 在 Elements/元素 标签页中查找：
   - `<section id="dualTableSection">` - 应该没有 `hidden` 属性
   - `<section id="singleTableSection" hidden>` - 应该有 `hidden` 属性
   - `<section id="statsSection" hidden>` - 应该有 `hidden` 属性
   - `<section id="previewSection" hidden>` - 应该有 `hidden` 属性

4. 点击"单表处理"按钮
5. 再次检查元素，应该看到：
   - `<section id="dualTableSection" hidden>` - 现在有 `hidden` 属性
   - `<section id="singleTableSection">` - 现在没有 `hidden` 属性

## 🎨 主题切换验证

### 双表模式（紫色）
- 背景渐变：淡紫色系
- 主按钮：紫色
- 激活状态：紫色高亮
- Body类应包含：`mode-dual`

### 单表模式（蓝色）
- 背景渐变：淡蓝色系
- 主按钮：蓝色
- 激活状态：蓝色高亮
- Body类应包含：`mode-single`

## ❌ 如果仍然看到问题

### 症状1: 两个模式区域同时显示
**可能原因**:
- JavaScript加载失败
- `switchMode()` 函数未执行
- CSS的 `[hidden]` 规则未生效

**解决步骤**:
1. 强制刷新: `Cmd+Shift+R` (Mac) 或 `Ctrl+Shift+R` (Windows)
2. 检查控制台是否有错误
3. 运行 `testUISeparation()` 查看状态

### 症状2: 共用区域始终显示
**可能原因**:
- `hideSharedSections()` 未在初始化时调用
- `hidden` 属性被其他代码移除

**解决步骤**:
1. 检查控制台运行: `hideSharedSections()`
2. 再次运行: `testUISeparation()`

### 症状3: 切换模式无反应
**可能原因**:
- 事件监听器未绑定
- 按钮选择器错误

**解决步骤**:
1. 控制台运行: `typeof dualTableMode`
2. 应该返回 `object`
3. 手动运行: `switchMode('single')`

## 📋 完整的DOM结构

```
body
  └── main.container
      ├── section.mode-selector-card (始终显示)
      ├── section#dualTableSection (仅双表模式)
      ├── section#singleTableSection (仅单表模式)
      ├── section#dataOperations.shared-section (有数据时)
      ├── section#statsSection.shared-section (有数据时)
      ├── section#previewSection.shared-section (有数据时)
      ├── section#historySection.shared-section (有历史时)
      └── section#templatesSection.shared-section (有模板时)
```

## 🔗 相关文件

- `index.html` - HTML结构
- `script.js` - JavaScript逻辑
- `styles.css` - CSS样式
- `UI_STRUCTURE.md` - 详细结构文档
- `TROUBLESHOOTING.md` - 故障排除指南

