#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from chinese_calendar import is_workday, is_holiday
import os
import math

# 设置显示中文
pd.set_option('display.unicode.east_asian_width', True)

# 默认配置
DEFAULT_CONFIG = {
    "files": {
        "staff_file": "人员.xlsx",
        "schedule_output": "排班表.xlsx",
        "config_file": "配置.xlsx"
    },
    "dates": {
        "start_date": "2025-04-01",
        "end_date": "2025-04-30"
    },
    "groups": {
        "daily_group_size": 5,
        "supplement_group_size": 2
    }
}


def load_config_from_excel(config_file=None):
    """
    从Excel文件加载配置，如果文件不存在则创建默认配置
    """
    config = DEFAULT_CONFIG.copy()

    # 如果指定了配置文件且文件存在，则加载配置
    if config_file and os.path.exists(config_file):
        try:
            # 读取文件路径配置
            files_df = pd.read_excel(config_file, sheet_name='文件配置')
            if not files_df.empty:
                for _, row in files_df.iterrows():
                    if '配置项' in row and '值' in row:
                        key = row['配置项']
                        value = row['值']
                        if key in config["files"] and isinstance(value, str):
                            config["files"][key] = value

            # 读取日期配置
            dates_df = pd.read_excel(config_file, sheet_name='日期配置')
            if not dates_df.empty:
                for _, row in dates_df.iterrows():
                    if '配置项' in row and '值' in row:
                        key = row['配置项']
                        value = row['值']
                        if key in config["dates"]:
                            if isinstance(value, str):
                                config["dates"][key] = value
                            elif isinstance(value, datetime):
                                config["dates"][key] = value.strftime('%Y-%m-%d')

            # 读取组大小配置
            groups_df = pd.read_excel(config_file, sheet_name='组配置')
            if not groups_df.empty:
                for _, row in groups_df.iterrows():
                    if '配置项' in row and '值' in row:
                        key = row['配置项']
                        value = row['值']
                        if key in config["groups"] and isinstance(value, (int, float)):
                            config["groups"][key] = int(value)

            print(f"已加载配置文件: {config_file}")
        except Exception as e:
            print(f"加载配置文件失败: {e}，将使用默认配置")
            # 如果配置文件加载失败，创建一个新的配置文件
            save_config_to_excel(config, config_file)
    else:
        # 如果配置文件不存在，创建一个默认配置文件
        save_config_to_excel(config, config_file if config_file else DEFAULT_CONFIG["files"]["config_file"])

    return config


def save_config_to_excel(config, config_file):
    """
    保存配置到Excel文件
    """
    try:
        with pd.ExcelWriter(config_file, engine='openpyxl') as writer:
            # 创建文件配置工作表数据
            files_data = [{"配置项": key, "值": value, "说明": get_config_description(key)}
                          for key, value in config.get("files", {}).items()]
            files_df = pd.DataFrame(files_data)
            files_df.to_excel(writer, sheet_name='文件配置', index=False)

            # 创建日期配置工作表数据
            dates_data = [{"配置项": key, "值": value, "说明": get_config_description(key)}
                          for key, value in config.get("dates", {}).items()]
            dates_df = pd.DataFrame(dates_data)
            dates_df.to_excel(writer, sheet_name='日期配置', index=False)

            # 创建组配置工作表数据
            groups_data = [{"配置项": key, "值": value, "说明": get_config_description(key)}
                           for key, value in config.get("groups", {}).items()]
            groups_df = pd.DataFrame(groups_data)
            groups_df.to_excel(writer, sheet_name='组配置', index=False)

        print(f"配置已保存到: {config_file}")
    except Exception as e:
        print(f"保存配置文件失败: {e}")


def get_config_description(key):
    """
    获取配置项的说明文字
    """
    descriptions = {
        "staff_file": "员工信息文件路径",
        "schedule_output": "排班表输出文件路径",
        "config_file": "配置文件路径",
        "start_date": "排班开始日期 (格式: YYYY-MM-DD)",
        "end_date": "排班结束日期 (格式: YYYY-MM-DD)",
        "daily_group_size": "日常组人数",
        "supplement_group_size": "增补组人数"
    }
    return descriptions.get(key, "")


def create_work_calendar(start_date, end_date):
    """
    创建工作日历，标记工作日、周末和法定假期
    """
    print(f"创建从 {start_date} 到 {end_date} 的工作日历")
    # 创建日期范围
    dates = pd.date_range(start=start_date, end=end_date)
    calendar_df = pd.DataFrame({'日期': dates})

    # 添加星期几和日期格式化
    calendar_df['星期'] = calendar_df['日期'].dt.dayofweek
    calendar_df['星期名'] = calendar_df['星期'].apply(
        lambda x: ['周一', '周二', '周三', '周四', '周五', '周六', '周日'][x])
    calendar_df['日期格式化'] = calendar_df['日期'].dt.strftime('%Y-%m-%d')

    # 标记是否为工作日
    calendar_df['是工作日'] = calendar_df['日期'].apply(lambda x: is_workday(x))

    # 标记是否为法定假日
    calendar_df['是法定假日'] = calendar_df['日期'].apply(lambda x: is_holiday(x))

    # 标记节假日后的第一个工作日
    calendar_df['节假日后第一天'] = False

    # 遍历日历，标记节假日或周末后的第一个工作日
    for i in range(1, len(calendar_df)):
        # 如果当天是工作日，且前一天不是工作日
        if calendar_df.loc[i, '是工作日'] and not calendar_df.loc[i - 1, '是工作日']:
            calendar_df.loc[i, '节假日后第一天'] = True

    # 过滤出工作日
    work_days = calendar_df[calendar_df['是工作日']].reset_index(drop=True)

    print(f"工作日历创建完成，包含 {len(work_days)} 个工作日")
    return work_days


def assign_daily_groups(staff_df, calendar_df, daily_group_size=5, supplement_group_size=2):
    """
    分配日常组和增补组 (v2 逻辑)
    - 每人每月最多排2次日常班.
    - 尽可能打散人员组合.
    - 增补组公平分配, 且尽量与日常组来自不同小组.
    """
    print("开始安排日常组和增补组 (v2 逻辑)")

    MAX_SHIFTS_PER_PERSON = 2

    # 为日历添加列
    calendar_df['日常组'] = None
    calendar_df['增补组'] = None

    # 复制员工数据用于更新排班过程中的计数
    staff_task_count = staff_df.copy()
    staff_task_count['上次排班周期'] = -100  # -100确保任何员工都可以被第一次选中
    staff_task_count['最近共事人员'] = [[] for _ in range(len(staff_task_count))]

    # 检查排班可行性
    total_periods = math.ceil(len(calendar_df) / 3)
    total_slots = total_periods * daily_group_size
    max_available_slots = len(staff_df) * MAX_SHIFTS_PER_PERSON

    if total_slots > max_available_slots:
        print(f"警告: 总排班需求 ({total_slots}人次) 超过了最大可排班次数 ({max_available_slots}人次)。")
        print("可能无法满足所有排班需求，或者部分人员排班会超过2次。请考虑调整排班周期或组大小。")

    # 按三天的周期进行排班
    current_period_idx = 0
    for period_idx in range(0, len(calendar_df), 3):
        current_period_idx += 1
        current_period = calendar_df.iloc[period_idx:period_idx + 3]

        # --- 1. 选择日常组 ---
        daily_group_members_df = pd.DataFrame()
        daily_team_groups_count = {}  # 用于跟踪每个小组已选入的人数

        # 获取候选人：日常次数小于最大限制
        candidates = staff_task_count[staff_task_count['日常次数'] < MAX_SHIFTS_PER_PERSON].copy()

        # 如果严格遵守规则后候选人不足，则放宽限制
        if len(candidates) < daily_group_size:
            print(f"警告: 周期 {current_period_idx}, 可用候选人 ({len(candidates)}) 不足 {daily_group_size}人。")
            if len(staff_task_count) >= daily_group_size:
                print("将从所有员工中选择，这可能会打破'每人最多排班2次'的限制。")
                candidates = staff_task_count.copy()
            else:
                print(f"错误: 总员工数 ({len(staff_task_count)}) 小于日常组大小 ({daily_group_size})，无法完成排班。")
                continue  # 跳过这个周期

        # 迭代选择组成员，直到满足人数要求 (v6: 同组人员不超过2个)
        daily_group_members_df = pd.DataFrame()
        for i in range(daily_group_size):
            if candidates.empty:
                print(f"警告: 在选择第 {i + 1} 名成员时，候选人池为空。")
                break

            # --- v6: 筛选出合格的候选人 ---
            # 排除那些来自"已满员"（超过2人）小组的候选人
            groups_at_limit = [group for group, count in daily_team_groups_count.items() if count >= 2]
            if groups_at_limit:
                eligible_candidates = candidates[~candidates['小组'].isin(groups_at_limit)]
            else:
                eligible_candidates = candidates

            if eligible_candidates.empty:
                print(f"警告: 周期 {current_period_idx}，在选择第 {i + 1} 名成员时，找不到符合同组不超过2人的候选人。")
                print("本周期的日常组将无法完全形成。")
                break

            # --- 计算每个合格候选人的得分 (v7 逻辑)---
            scores = pd.Series(0, index=eligible_candidates.index, dtype=float)

            # 规则a: 优先选择排班次数少的人 (权重最高)
            scores += (MAX_SHIFTS_PER_PERSON - eligible_candidates['日常次数']) * 100

            # 规则b: 优先选择距离上次排班时间久的人
            scores += (current_period_idx - eligible_candidates['上次排班周期']) * 10

            # 规则c (软惩罚): 尽可能避免选择第2名同组人员
            groups_with_one_member = [group for group, count in daily_team_groups_count.items() if count == 1]
            if groups_with_one_member:
                scores[eligible_candidates['小组'].isin(groups_with_one_member)] -= 50  # 对成为"第二人"的候选人进行分数惩罚
            
            # 规则d (惩罚项): 如果与已选成员近期共事过，则扣分
            if not daily_group_members_df.empty:
                current_members_names = daily_group_members_df['人员名称'].tolist()
                for cand_idx, cand_row in eligible_candidates.iterrows():
                    co_work_count = sum(1 for m in current_members_names if m in cand_row['最近共事人员'])
                    scores.loc[cand_idx] -= co_work_count * 20

            # 选取得分最高的候选人
            best_candidate_idx = scores.idxmax()
            best_candidate = eligible_candidates.loc[[best_candidate_idx]]

            # 将选中的人添加到日常组，并从 *总候选人池* 中移除
            daily_group_members_df = pd.concat([daily_group_members_df, best_candidate])
            candidates = candidates.drop(best_candidate_idx)

            # 更新小组计数
            selected_group = best_candidate.iloc[0]['小组']
            daily_team_groups_count[selected_group] = daily_team_groups_count.get(selected_group, 0) + 1
            
        daily_group_members = daily_group_members_df['人员名称'].tolist()

        # 如果人数不足，则不再继续（前面已有警告）
        if len(daily_group_members) < daily_group_size:
            print(f"错误: 未能为周期 {current_period_idx} 选出 {daily_group_size} 名日常组成员。")
            continue

        # --- 2. 为周期内的每一天分配日常组和增补组 ---
        for day_idx in current_period.index:
            calendar_df.loc[day_idx, '日常组'] = '; '.join(daily_group_members)

            # 如果是节假日后第一天，则需要安排增补组
            if calendar_df.loc[day_idx, '节假日后第一天']:
                # --- v5: 组合限制规则 ---
                supp_candidates_pool = staff_task_count[~staff_task_count['人员名称'].isin(daily_group_members)].copy()
                supplement_group_df = pd.DataFrame()
                
                # 获取日常组的组构成
                daily_team_groups = daily_group_members_df['小组'].value_counts().to_dict()

                for _ in range(supplement_group_size):
                    if supp_candidates_pool.empty:
                        print(f"警告: 增补组候选人不足，无法选出 {supplement_group_size} 人。")
                        break
                    
                    # 确定合格候选人
                    eligible_indices = []
                    # 获取已选增补人员的组构成
                    supp_team_groups = supplement_group_df['小组'].value_counts().to_dict() if not supplement_group_df.empty else {}

                    for idx, cand in supp_candidates_pool.iterrows():
                        cand_group = cand['小组']
                        # 计算如果选上该候选人，其所在小组的总人数
                        total_from_group = daily_team_groups.get(cand_group, 0) + supp_team_groups.get(cand_group, 0)
                        # 规则: 同组总人数不能超过2人
                        if total_from_group < 2:
                            eligible_indices.append(idx)
                    
                    # 优先从合格者中选，如果没有则打破规则
                    if eligible_indices:
                        selection_pool = supp_candidates_pool.loc[eligible_indices]
                    else:
                        print(f"警告: 为满足增补组人数，将打破'日常组+增补组同组人员不超过2人'的限制。")
                        selection_pool = supp_candidates_pool

                    # 从确定的池中，选择增补次数最少的一位
                    best_candidate = selection_pool.sort_values(by='增补次数', ascending=True).head(1)
                    
                    # 添加到增补组并从候选池中移除
                    supplement_group_df = pd.concat([supplement_group_df, best_candidate])
                    supp_candidates_pool = supp_candidates_pool.drop(best_candidate.index[0])

                # 更新增补组信息和次数统计
                if not supplement_group_df.empty:
                    supplement_group = supplement_group_df['人员名称'].tolist()
                    calendar_df.loc[day_idx, '增补组'] = '; '.join(supplement_group)

                    for name in supplement_group:
                        staff_idx = staff_task_count.index[staff_task_count['人员名称'] == name]
                        staff_task_count.loc[staff_idx, '增补次数'] += 1
                        staff_task_count.loc[staff_idx, '节假日后一天'] += 1

        # --- 3. 更新本次日常组成员的统计数据 ---
        for name in daily_group_members:
            staff_idx = staff_task_count.index[staff_task_count['人员名称'] == name][0]
            staff_task_count.loc[staff_idx, '日常次数'] += 1
            staff_task_count.loc[staff_idx, '上次排班周期'] = current_period_idx

            # 更新共事记录
            current_coworkers = staff_task_count.loc[staff_idx, '最近共事人员']
            for co_member in daily_group_members:
                if name != co_member and co_member not in current_coworkers:
                    current_coworkers.append(co_member)
            staff_task_count.at[staff_idx, '最近共事人员'] = current_coworkers

    # --- 最终检查与报告 ---
    print("\n排班完成，最终统计:")
    max_daily = staff_task_count['日常次数'].max()
    min_daily = staff_task_count['日常次数'].min()
    print(f"日常排班次数范围: {min_daily} - {max_daily}")
    if max_daily - min_daily > 1:
        print(f"警告: 日常排班次数差距为 {max_daily-min_daily}，大于1。可考虑重新引入平衡函数。")

    print("\n各人员排班次数详情:")
    for _, row in staff_task_count.sort_values(by=['小组', '人员名称']).iterrows():
        print(f"- {row['人员名称']} (小组: {row['小组']}): 日常 {row['日常次数']}次, 增补 {row['增补次数']}次")

    return calendar_df, staff_task_count


def export_schedule(calendar_df, output_file, daily_group_size, supplement_group_size):
    """
    导出排班表到Excel，将每位成员放在单独的单元格中。
    """
    print(f"导出格式化排班表: {output_file}")
    
    # 基础信息
    schedule_export_df = calendar_df[['日期格式化', '星期名', '是法定假日', '节假日后第一天']].copy()
    schedule_export_df.rename(columns={'日期格式化': '日期', '星期名': '星期'}, inplace=True)
    
    # 创建空的日常组和增补组列
    for i in range(1, daily_group_size + 1):
        schedule_export_df[f'日常组{i}'] = ''
    for i in range(1, supplement_group_size + 1):
        schedule_export_df[f'增补组{i}'] = ''
        
    # 填充数据
    for index, row in calendar_df.iterrows():
        # 填充日常组
        if pd.notna(row['日常组']):
            daily_members = row['日常组'].split('; ')
            for i, member in enumerate(daily_members):
                if i < daily_group_size:
                    schedule_export_df.loc[index, f'日常组{i+1}'] = member
        
        # 填充增补组
        if pd.notna(row['增补组']):
            supp_members = row['增补组'].split('; ')
            for i, member in enumerate(supp_members):
                if i < supplement_group_size:
                    schedule_export_df.loc[index, f'增补组{i+1}'] = member

    schedule_export_df.to_excel(output_file, index=False)
    print(f"排班表已导出到: {output_file}")


def export_updated_staff(staff_df, output_file):
    """
    导出更新后的人员信息
    """
    print(f"导出更新后的人员信息: {output_file}")
    staff_df.to_excel(output_file, index=False)
    print(f"更新后的人员信息已导出到: {output_file}")
    return staff_df


def main():
    """
    主函数
    """
    print("开始执行排班系统...")

    # 加载配置文件（默认从"配置.xlsx"加载）
    config_file = "配置.xlsx"
    config = load_config_from_excel(config_file)

    # 提取配置参数
    staff_file = config["files"]["staff_file"]
    schedule_output = config["files"]["schedule_output"]
    start_date = config["dates"]["start_date"]
    end_date = config["dates"]["end_date"]
    daily_group_size = config["groups"]["daily_group_size"]
    supplement_group_size = config["groups"]["supplement_group_size"]

    print(f"\n当前配置:")
    print(f"员工信息文件: {staff_file}")
    print(f"排班表输出: {schedule_output}")
    print(f"排班日期: {start_date} 至 {end_date}")
    print(f"日常组人数: {daily_group_size}")
    print(f"增补组人数: {supplement_group_size}\n")

    # --- 1. 加载和筛选员工数据 ---
    print(f"读取员工信息表: {staff_file}")
    try:
        staff_df_full = pd.read_excel(staff_file)
    except FileNotFoundError:
        print(f"错误: 员工信息文件未找到于 '{staff_file}'")
        print("请确保文件名和路径正确，或通过 '配置.xlsx' 文件进行配置。")
        return

    # 为后续合并，保留原始 DataFrame 的副本
    staff_df_to_update = staff_df_full.copy()
    
    # 筛选参与排班的人员
    if '是否参与排班' in staff_df_to_update.columns:
        initial_count = len(staff_df_to_update)
        # 根据用户新需求：排除"是否参与排班"列中 *有任何值* 的人员
        staff_df_for_scheduling = staff_df_to_update[staff_df_to_update['是否参与排班'].isnull()].copy()
        print(f"已根据 '是否参与排班' 列进行筛选。仅该列为空白的人员参与排班。")
        print(f"参与排班人数: {len(staff_df_for_scheduling)} (总人数: {initial_count})")
    else:
        staff_df_for_scheduling = staff_df_to_update.copy()
        print("提示: 未在员工表中找到 '是否参与排班' 列，默认所有人员都参与排班。")
        print("如需排除，请添加 '是否参与排班' 列，并在不参与排班人员的对应单元格中填写任意值。")

    # 确保所需列存在于将要用于排班的 DataFrame 中
    for col in ['日常次数', '增补次数', '节假日后一天']:
        if col not in staff_df_for_scheduling.columns:
            staff_df_for_scheduling[col] = 0

    print(f"成功读取并处理员工数据，共 {len(staff_df_for_scheduling)} 名员工参与排班。")

    if staff_df_for_scheduling.empty:
        print("没有员工参与排班，程序终止。")
        return

    # --- 2. 创建日历和分配任务 ---
    work_calendar = create_work_calendar(start_date, end_date)

    schedule_df, updated_staff_counts = assign_daily_groups(
        staff_df_for_scheduling,
        work_calendar,
        daily_group_size=daily_group_size,
        supplement_group_size=supplement_group_size
    )

    # --- 3. 导出排班表 ---
    export_schedule(schedule_df, schedule_output, daily_group_size, supplement_group_size)

    # --- 4. 更新原始员工表中的排班次数并保存 ---
    print(f"\n正在更新原始员工文件: {staff_file}")

    # 使用 '人员名称' 作为键，将更新后的次数合并回原始的DataFrame副本
    staff_df_to_update.set_index('人员名称', inplace=True)
    updated_staff_counts.set_index('人员名称', inplace=True)

    # 更新 staff_df_to_update 中存在的员工的次数
    staff_df_to_update.update(updated_staff_counts[['日常次数', '增补次数', '节假日后一天']])

    # 恢复索引
    staff_df_to_update.reset_index(inplace=True)
    
    # 将完全更新后的DataFrame保存回 *原始* 的员工文件
    export_updated_staff(staff_df_to_update, staff_file)

    print("\n排班系统执行完成！")
    print(f"1. 排班表已导出到: {schedule_output}")
    print(f"2. 原始员工信息文件 '{staff_file}' 已更新排班次数。")


if __name__ == "__main__":
    main()
