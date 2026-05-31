# UI for Self

个人 UI 特效与组件合集，收集自用的小工具、特效组件和实验性页面。

## 组件

| 文件/目录 | 说明 | 预览 |
|-----------|------|------|
| [radix-color-browser.html](radix-color-browser.html) | Radix Colors 可视化色板浏览器，31色×12级，亮暗切换+选色+复制 | 双击打开 |
| [fluid-bg/](fluid-bg/) | WebGL 流体模拟背景，鼠标拖拽产生流动效果 | [demo](fluid-bg/demo.html) |

---

## radix-color-browser

基于 [Radix Colors](https://www.radix-ui.com/colors) 的交互式色板浏览器。

- 31 个色彩家族，每个 12 级色阶
- 亮/暗模式一键切换
- 点击色块选中，底部显示已选列表
- 一键复制选色结果（如 `blue 9, slate 12, amber 3`）
- 离线可用，零依赖，单文件 HTML

## fluid-bg

基于 [PavelDoGreat/WebGL-Fluid-Simulation](https://github.com/PavelDoGreat/WebGL-Fluid-Simulation) (MIT) 的流体背景组件。

- 一行引入：`<script src="fluid-bg.js"></script>`
- 支持自定义参数（颜色强度、拖尾力度、涡流、Bloom 等）
- 打开 `demo.html` 即可体验并调节参数

```html
<script>
  window.FluidBg = {
    SPLAT_INTENSITY: 0.08,
    SPLAT_FORCE: 4000,
    CURL: 25
  };
</script>
<script src="fluid-bg.js"></script>
```

---

## Design Studio（交互式设计流水线）

把选色、选风格、出成品串成一条流水线，配合 Claude Code 使用。

### 流程

```
选色 (color-picker.html)
  → 选风格 (style-picker.html)
  → 品味约束 (taste-picker.html)
  → 告诉 Claude 内容需求
  → Claude 调用 skill 出成品
```

### 用法

1. 打开 `design-studio/color-picker.html`，选 3-5 个颜色，点「确认选择」→ 复制 JSON
2. 回到 Claude Code 粘贴选色结果
3. Claude 生成带选色的风格选择页 → 打开选风格 → 复制 JSON
4. 打开 `taste-picker.html`（带 `?colors=...&style=...` 参数），调整旋钮和规则，点「确认输出」→ 复制 JSON
5. 回到 Claude Code 粘贴，然后告诉 Claude 你要做什么内容
6. Claude 生成最终 HTML/PPT

### 文件

| 文件 | 说明 |
|------|------|
| [color-picker.html](design-studio/color-picker.html) | Step 1: Radix Colors 取色器，选色后导出 JSON |
| [style-picker.html](design-studio/style-picker.html) | Step 2: 风格选择器，PPT 模板 + 网页品牌风格 |
| [taste-picker.html](design-studio/taste-picker.html) | Step 3: 品味约束，反 AI 味规则 + 三旋钮调参，导出约束 JSON |

---

## 约定

- 每个组件一个目录或单文件，自包含，拖出来即可使用
- 组件名用功能名：`fluid-bg`、`glass-card`、`gradient-text`
- 单文件组件直接放根目录，复杂组件放子目录并带 `demo.html`

## License

各组件沿用其上游开源协议，具体见各目录下的说明。
