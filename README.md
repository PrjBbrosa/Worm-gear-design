
# Light Worm Gear Tool v3 — 蜗轮蜗杆详细版（可扩展）

- 几何页：带示意图（帮助新手理解参数）
- 材料页：蜗杆材料默认 37CrS4（JSON库可导入）；蜗轮材料提供 PA66 draft，可编辑 E(T) 与 SN
- 曲线页：接触应力/齿根应力/输出扭矩波动/效率与接触数代理
- 疲劳页：雨流计数 + Miner 损伤（基于齿根应力代理）
- 导出：XLSX（整周期曲线）

## 运行
```bash
pip install -r requirements.txt
python app.py
```

> 当前接触应力与齿根应力为“轻量代理模型”，用于趋势与方案比选；后续可替换为更严谨的ISO/AGMA/数值接触模型。
