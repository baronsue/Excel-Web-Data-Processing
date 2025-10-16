# UI 分离问题排查指南

## 🔍 问题诊断步骤

### 1. 打开浏览器开发者控制台
- Chrome/Edge: `F12` 或 `Cmd+Option+I` (Mac) / `Ctrl+Shift+I` (Windows)
- Safari: `Cmd+Option+C` (需先在偏好设置中启用开发菜单)

### 2. 在控制台运行测试命令
```javascript
testUISeparation()
```

### 3. 检查输出

#### 期望的输出（初始状态 - 双表模式，无数据）:
```
=== UI 分离状态检查 ===
当前模式: dual
Body类: mode-dual

模式区域:
  双表区域 hidden: false    ✅ 应该显示
  单表区域 hidden: true     ✅ 应该隐藏

共用区域:
  统计区域 hidden: true     ✅ 无数据时隐藏
  预览区域 hidden: true     ✅ 无数据时隐藏
  数据操作 hidden: true     ✅ 无数据时隐藏

数据状态:
  result.rows: 0
  processedData.rows: 0
```

#### 切换到单表模式后（点击"单表处理"按钮）:
```
=== UI 分离状态检查 ===
当前模式: single
Body类: mode-single

模式区域:
  双表区域 hidden: true     ✅ 应该隐藏
  单表区域 hidden: false    ✅ 应该显示

共用区域:
  统计区域 hidden: true     ✅ 无数据时隐藏
  预览区域 hidden: true     ✅ 无数据时隐藏
  数据操作 hidden: true     ✅ 无数据时隐藏
```

## 🔧 手动测试命令

### 强制切换到双表模式
```javascript
switchMode('dual')
testUISeparation()
```

### 强制切换到单表模式
```javascript
switchMode('single')
testUISeparation()
```

### 手动显示共用区域
```javascript
showSharedSections()
testUISeparation()
```

### 手动隐藏共用区域
```javascript
hideSharedSections()
testUISeparation()
```

## ❌ 常见问题

### 问题1: 两个模式区域都显示
**原因**: `hidden` 属性没有正确设置
**解决**: 检查 `toggleVisibility` 函数是否正常工作

### 问题2: 共用区域始终显示
**原因**: `hidden` 属性没有被设置
**解决**: 检查 `hideSharedSections()` 是否被调用

### 问题3: 切换模式没反应
**原因**: 事件绑定失败或函数未定义
**解决**: 检查控制台是否有JavaScript错误

### 问题4: 主题颜色没有切换
**原因**: body类没有正确添加
**解决**: 检查 `document.body.className` 是否包含 `mode-dual` 或 `mode-single`

## 🐛 调试清单

- [ ] 刷新浏览器 (Cmd+Shift+R / Ctrl+Shift+R 强制刷新)
- [ ] 清除浏览器缓存
- [ ] 检查控制台是否有JavaScript错误
- [ ] 运行 `testUISeparation()` 查看当前状态
- [ ] 尝试手动切换模式
- [ ] 检查HTML元素是否存在
- [ ] 验证CSS文件已加载

## 📞 如果问题仍然存在

请在控制台运行以下命令并提供输出:

```javascript
console.log('jQuery $:', typeof $);
console.log('dualTableMode:', dualTableMode);
console.log('singleTableMode:', singleTableMode);
console.log('switchMode:', typeof switchMode);
testUISeparation();
```

