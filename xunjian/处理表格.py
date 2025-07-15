import os
import pandas as pd
from datetime import datetime

# 获取当前工作目录
FOLDER_PATH = os.getcwd()

# 定义文件路径
INSPECT_FILE = os.path.join(FOLDER_PATH, '巡检数据.xlsx')  # 巡检数据文件路径
OUTPUT_FILE = os.path.join(FOLDER_PATH, 'output.xlsx')  # 输出文件路径
LOG_FILE = os.path.join(FOLDER_PATH, 'deleted_records.csv')  # 删除记录日志文件

def convert_to_float(value):
    """将值转换为浮点数，空值或无效值返回0"""
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0

def remove_duplicates(df, key_column, date_column):
    """
    删除重复项，只保留指定日期最新的记录
    若有最新的时间重复的也保留
    """
    if key_column not in df.columns or date_column not in df.columns:
        print(f"未找到'{key_column}'或'{date_column}'列，无法删除重复项")
        return df

    # 临时将日期转换为datetime进行比较
    df[date_column + '_temp'] = pd.to_datetime(df[date_column], errors='coerce')

    # 按关键列分组，找到每个关键列的最大日期
    max_dates = df.groupby(key_column)[date_column + '_temp'].transform('max')

    # 保留日期等于最大日期的记录
    df_filtered = df[df[date_column + '_temp'] == max_dates].copy()

    # 记录删除的重复项
    df_deleted = df[df[date_column + '_temp'] < max_dates].copy()

    # 保存删除的记录到日志文件
    if not df_deleted.empty:
        print(f"正在保存{len(df_deleted)}条删除记录到: {LOG_FILE}")
        df_deleted.to_csv(LOG_FILE, index=False)

    # 删除临时列，不保留在最终结果中
    df_filtered = df_filtered.drop(columns=[date_column + '_temp'])

    return df_filtered


def create_output_table(df_inspect):
    """创建输出表格，包含指定字段并处理运营商数据"""

    # 定义输出表格的列名
    output_columns = [
        '县区', '站址名称', '站址编码', '站址运维ID', '塔型（资源数据）',
        '塔高（资源数据）', '运营商共享情况（运监数据）', '巡检人员',
        '巡检日期', '审批编号', '移动_挂高（米）', '移动_在用天线数',
        '移动_闲置天线数', '移动_RRU数（台）', '移动_AAU数（台）',
        '联通_挂高（米）', '联通_在用天线数', '联通_闲置天线数',
        '联通_RRU数（台）', '联通_AAU数（台）', '电信_挂高（米）',
        '电信_在用天线数', '电信_闲置天线数', '电信_RRU数（台）',
        '电信_AAU数（台）'
    ]

    # 创建空的DataFrame用于输出
    df_output = pd.DataFrame(columns=output_columns)

    # 从原始数据中复制基础信息
    base_columns = [
        '县区', '站址名称', '站址编码', '站址运维ID', '塔型（资源数据）',
        '塔高（资源数据）', '运营商共享情况（运监数据）', '巡检人员',
        '巡检日期', '审批编号'
    ]

    # 确保基础列存在于原始数据中
    available_base_columns = [col for col in base_columns if col in df_inspect.columns]
    df_output[available_base_columns] = df_inspect[available_base_columns].copy()

    # 处理运营商数据
    operators = ['移动', '联通', '电信']
    attributes = ['挂高（米）', '在用天线数', '闲置天线数', 'RRU数（台）', 'AAU数（台）']

    for operator in operators:
        # 筛选当前运营商的数据
        operator_mask = df_inspect['运营商'].str.contains(operator, na=False)
        operator_data = df_inspect[operator_mask].copy()

        if not operator_data.empty:
            # 按站址编码分组处理
            grouped = operator_data.groupby('站址编码')

            for site_code, group in grouped:
                # 找到输出表中对应的行
                output_idx = df_output[df_output['站址编码'] == site_code].index

                if not output_idx.empty:
                    # 挂高取最大值
                    height_col = f'{operator}_挂高（米）'
                    if '挂高（米）' in group.columns:
                        df_output.loc[output_idx, height_col] = group['挂高（米）'].apply(convert_to_float).max()

                    # 其他属性取和
                    for attr in ['在用天线数', '闲置天线数', 'RRU数（台）', 'AAU数（台）']:
                        output_col = f'{operator}_{attr}'
                        if attr in group.columns:
                            df_output.loc[output_idx, output_col] = group[attr].apply(convert_to_float).sum()

    # 填充NaN值为0
    operator_columns = [col for col in output_columns if col not in base_columns]
    df_output[operator_columns] = df_output[operator_columns].fillna(0)

    return df_output


def remove_output_duplicates(df_output):
    """删除输出表中站址编码和审批编号完全重复的记录"""
    if '站址编码' not in df_output.columns or '审批编号' not in df_output.columns:
        print("缺少'站址编码'或'审批编号'列，无法进行去重")
        return df_output

    # 记录去重前的记录数
    before_count = len(df_output)

    # 删除完全重复的记录（基于站址编码和审批编号）
    df_output = df_output.drop_duplicates(subset=['站址编码', '审批编号'], keep='first')

    # 记录去重后的记录数
    after_count = len(df_output)

    if before_count > after_count:
        print(f"删除了 {before_count - after_count} 条重复记录（站址编码+审批编号重复）")

    return df_output


def process_files():
    try:
        # 检查巡检数据文件和输出表格模板文件是否存在
        if not os.path.exists(INSPECT_FILE):
            print(f"巡检数据文件 {INSPECT_FILE} 不存在")
            return

        # 读取巡检数据文件，指定表头在第二行
        print(f"正在读取巡检数据文件: {INSPECT_FILE}")
        df_inspect = pd.read_excel(INSPECT_FILE, header=1, dtype=str)  # 修改点：添加header=1参数

        # 删除 '巡检数据' 表中的重复项
        print("正在处理 '巡检数据' 表中的重复项...")
        df_inspect = remove_duplicates(df_inspect, '站址编码', '巡检日期')

        # 创建输出表格
        print("正在创建输出表格...")
        df_output = create_output_table(df_inspect)

        # 删除输出表中的重复项（站址编码+审批编号）
        print("正在处理输出表中的重复项...")
        df_output = remove_output_duplicates(df_output)

        # 保存到输出文件
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            df_output.to_excel(writer, sheet_name='巡检汇总', index=False)

        print(f"数据处理完成，结果已保存到: {OUTPUT_FILE}")
        if os.path.exists(LOG_FILE):
            print(f"删除记录已保存到: {LOG_FILE}")

    except Exception as e:
        print(f"处理过程中出错: {str(e)}")
        raise


if __name__ == "__main__":
    process_files()