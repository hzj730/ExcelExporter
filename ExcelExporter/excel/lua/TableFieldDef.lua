-- 该文件自动生成，请不要随修改
local fields = {}
fields.DefaultNum = 0
fields.DefaultStr = ""
fields.DefaultTable = {}

-- 所有lua表的字段定义
fields.TableDefine = {
activity_summary = {
    meta = {
        id = 1, -- 键名
        activity_name = 2, -- 活动名
        disable = 3, -- 禁用
        param_reward = 4, -- 奖励参数
        param_show = 5, -- 展示参数
        preview_start_ts = 6, -- 预告起始时间戳
        start_ts_type = 7, -- 起始时间类型
        start_ts = 8, -- 起始时间戳
        end_ts = 9, -- 结束时间戳
        duration = 10, -- 持续时间
    },
    file = 'activity_summary.lua',
},
daily_sign_in = {
    meta = {
        id = 1, -- 唯一id
        reward = 2, -- 奖励
    },
    file = 'daily_sign_in.lua',
},
new_account = {
    meta = {
        id = 1, -- 唯一id
        quest_name = 2, -- 任务描述
        quest_set = 3, -- 任务所在组
        objective_num = 4, -- 要求数量
        objective_data = 5, -- 目标数据
        goto_feature = 6, -- 前往功能
        reward = 7, -- 完成后奖励
        quest_point = 8, -- 完成积分
    },
    file = 'new_account.lua',
},
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
