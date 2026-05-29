# UI Lab

个人 UI 特效与组件合集，收集自用的小工具、特效组件和实验性页面。

## 组件

| 目录 | 说明 | 预览 |
|------|------|------|
| [fluid-bg/](fluid-bg/) | WebGL 流体模拟背景，鼠标拖拽产生流动效果 | [demo](fluid-bg/demo.html) |

---

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

## 约定

- 每个组件一个目录，自包含，拖出来即可使用
- 每个组件都有 `demo.html`，双击直接看效果
- 组件名用功能名：`fluid-bg`、`glass-card`、`gradient-text`

## License

各组件沿用其上游开源协议，具体见各目录下的说明。
