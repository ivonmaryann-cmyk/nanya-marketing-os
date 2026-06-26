# 测试脚本说明

请把 `tests/` 目录复制到安装后的项目根目录下运行。

## 1. 不依赖 Excel 的静态测试

```powershell
python tests\smoke_order_reprice.py
```

这个测试只检查模块导入和核心 helper，不需要真实客户数据。

## 2. 使用样例 Excel 的回归测试

```powershell
$env:ORDER_REPRICE_SAMPLE_ROOT="D:\桌面\胜宏"
python tests\test_with_sample_paths.py
```

目录结构建议：

```text
D:\桌面\胜宏
  功能一原始数据
  功能二原始数据
  功能三原始数据
```

如果没有样例 Excel，脚本会输出 SKIP。

