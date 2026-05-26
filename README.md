# UI for Self

个人 UI 效果合集，收集自用的小工具、特效组件和实验性页面。

## 目录

| 目录 | 说明 | 预览 |
|------|------|------|
| [fluid-bg/](fluid-bg/) | WebGL 流体模拟背景，鼠标拖拽即可产生流动效果 | — |

---

## fluid-bg

基于 [PavelDoGreat/WebGL-Fluid-Simulation](https://github.com/PavelDoGreat/WebGL-Fluid-Simulation) (MIT) 的流体背景组件。

**特点：**

- 一行引入：`<script src="fluid-bg.js"></script>`
- 支持自定义参数（颜色强度、拖尾力度、涡流、Bloom 等）
- 鼠标拖拽自动产生流体效果
- 打开 `demo.html` 即可体验并调节参数

**快速使用：**

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

## License

各组件沿用其上游开源协议，具体见各目录下的说明。
