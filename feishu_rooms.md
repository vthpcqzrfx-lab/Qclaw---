# 飞书会议室清单

> 中电金信大厦 - 北京
> 同步时间：2026-04-02
> 共收录 27 个会议室

## 北楼4F（17个）

| 会议室名称 | 容量 | room_id |
|-----------|------|---------|
| 安全可靠 | 8人 | `omm_b35390646be430475d29077ea59102bb` |
| 机器学习 | 8人 | `omm_9f1cc4740e02dab5fa36d0486c417cac` |
| 个性学习 | 8人 | `omm_3d74c3bd150c80b1f983e5e3c5fcc378` |
| 人工智能 | 8人 | `omm_36d7da02bd402dcec702b9a8111a5586` |
| 代码之美 | 8人 | `omm_ec54569e2ecc752b1e3a2ca3ce728d85` |
| 双机热备 | 8人 | `omm_f7e4d1d352785e9610e31e6742ea76e5` |
| 头脑风暴 | 8人 | `omm_5b2a67ca0b32537abd089923b71899e6` |
| 数据结构 | 8人 | `omm_f8393fad4cb86aaca3be1902cbb5563f` |
| 数据至上 | 8人 | `omm_5a594599fb707dca9fb774a56529edb9` |
| 数据驱动 | 8人 | `omm_60c2bf739aae8df1e29e7a536153573c` |
| 架构之美 | 8人 | `omm_98a70876faae6460a9c9517ef125ed9a` |
| 海量存储 | 8人 | `omm_a47ef324b4f944619b94fefd2260d6c8` |
| 用户思维 | 8人 | `omm_d5413d7f29ded22aa22cf1d47c68db33` |
| 稳定高效 | 30人 | `omm_127b8c11f087a32e6b79d708fc928f12` |
| 算法之美 | 8人 | `omm_ebbd0c860daef75d584b8ce396382900` |
| 算法导论 | 8人 | `omm_40f4e7ae694cc9aa987cf1f5f85b1407` |
| 面向对象 | 8人 | `omm_21b0c7018103e12f443f8970f353af30` |

## 南楼2F（10个）

| 会议室名称 | 容量 | room_id |
|-----------|------|---------|
| 公正 | 8人 | `omm_42e3817e34d6e588504a903ca0ad8cbe` |
| 务实 | 8人 | `omm_2fa7e0a03e30d47025240bd51e5912d8` |
| 友善 | 8人 | `omm_02153bbef8c9b309f67145554530a441` |
| 合作 | 8人 | `omm_4c1b0f2bb1a02d024625a5829a95afd8` |
| 敬业 | 10人 | `omm_f8558f65d9a48bb745eb756b18e00cbe` |
| 文明 | 8人 | `omm_5795c988e27afed22b53e43137446905` |
| 自由 | 20人 | `omm_dac321fa1b45141b7256f90d72649fe2` |
| 诚信 | 8人 | `omm_db8ce110a8e9b4877dc9bb3e59f778a7` |
| 进取 | 8人 | `omm_230cfffa786337f955c501a223c8f554` |

## 使用说明

### 预约流程
1. 创建日历事件
2. 添加会议室作为attendee（type: resource, room_id: xxx）
3. 查询rsvp_status确认是否成功：
   - `accept` - 预约成功 ✅
   - `decline` - 已被占用 ❌
   - `needs_action` - 待确认 ⏳

### 自动抢占策略
- 遍历所有会议室尝试预约
- 检测rsvp_status找到accept状态的会议室
- 返回第一个可用的会议室

### API凭证
- app_id: `cli_a940eaf4bd389bcb`
- app_secret: `GSyBCdMXFoqHVqvs7FoQT7IdPQfmnbb7`
