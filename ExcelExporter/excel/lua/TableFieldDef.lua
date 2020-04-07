-- 该文件自动生成，请不要随修改
local fields = {}
fields.DefaultNum = 0
fields.DefaultStr = ""
fields.DefaultTable = {}

-- 所有lua表的字段定义
fields.TableDefine = {
abyss = {
    meta = {
        id = 1, -- 唯一id
        type = 2, -- 深渊类型
        level = 3, -- 层数
        next_id = 4, -- 下一关
        banned = 5, -- 阵容限制
        bg_img = 6, -- 关卡背景图
        clear_extra_reward = 7, -- 额外通关奖励
        boss_flag = 8, -- Boss标志
        level_enemies_config = 9, -- 关卡敌人配置
    },
    file = 'abyss.lua',
},
Sample = {
    meta = {
        id = 1, -- 唯一id
        type = 2, -- 名称
        fom1 = 3, -- 公式1
        next_id = 4, -- 下一关
        banned = 5, -- 阵容限制
        bg_img = 6, -- 关卡背景图
        clear_extra_reward = 7, -- 额外通关奖励
        boss_flag = 8, -- Boss标志
        desc = 9, -- 关卡敌人配置
    },
    file = 'Sample.lua',
},
}

_G.TableField = fields
return _G.TableField
