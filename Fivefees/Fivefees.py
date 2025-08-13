import os
import glob
import pythoncom
import win32com.client as win32
import time


class FeeSummarizer:
    def __init__(self):
        INDEX = os.getcwd()
        self.xls_path = os.path.join(INDEX, "xls")
        self.model_path = os.path.join(self.xls_path, "model")
        self.output_path = os.path.join(INDEX, "output")
        self.city_path = os.path.join(self.xls_path, "city_data")

        # 创建输出目录（如果不存在）
        os.makedirs(self.output_path, exist_ok=True)

        # 定义每种费用类型的配置，包含需要设置文本格式的列
        self.fee_configs = {
            '电费': {
                'model_path': os.path.join(self.model_path, "附件1： 电费稽核清单.xlsx"),
                'output_path': os.path.join(self.output_path, "附件1： 电费稽核清单.xlsx"),
                'source_sheet_names': ['2-3电费稽核问题清单模板'],
                'target_sheet_index': 1,
                'data_range': 'A:AV',
                'start_row': 4,
                'file_keywords': ['电费'],
                'text_columns': ['D', 'G']  # 电费表需要文本格式的列
            },
            '场地费': {
                'model_path': os.path.join(self.model_path, "附件2：场地费稽核清单.xlsx"),
                'output_path': os.path.join(self.output_path, "附件2：场地费稽核清单.xlsx"),
                'source_sheet_names': ['3-3场地费稽核问题清单模板'],
                'target_sheet_index': 1,
                'data_range': 'A:AG',
                'start_row': 4,
                'file_keywords': ['场地费'],
                'text_columns': ['D']  # 场地费表需要文本格式的列
            },
            '修理维护费': {
                'model_path': os.path.join(self.model_path, "附件3：修理维护费稽核清单.xlsx"),
                'output_path': os.path.join(self.output_path, "附件3：修理维护费稽核清单.xlsx"),
                'subtypes': [
                    {
                        'name': '维护费',
                        'source_sheet_name': ['4-3-1维护费稽核问题清单模板'],
                        'target_sheet_index': 1,
                        'data_range': 'A:X',
                        'start_row': 4,
                        'file_keywords': ['维护费'],
                        'text_columns': ['D']  # 维护费表需要文本格式的列
                    },
                    {
                        'name': '修理费',
                        'source_sheet_name': ['4-3-2修理费稽核问题清单模板'],
                        'target_sheet_index': 2,
                        'data_range': 'A:AC',
                        'start_row': 4,
                        'file_keywords': ['维护费'],
                        'text_columns': ['F']  # 修理费表需要文本格式的列
                    }
                ]
            },
            '发电费': {
                'model_path': os.path.join(self.model_path, "附件4：发电费稽核清单.xlsx"),
                'output_path': os.path.join(self.output_path, "附件4：发电费稽核清单.xlsx"),
                'source_sheet_names': ['5-3发电费稽核问题清单模板'],
                'target_sheet_index': 1,
                'data_range': 'A:AR',
                'start_row': 4,
                'file_keywords': ['发电费'],
                'text_columns': ['F', 'G']  # 发电费表需要文本格式的列
            },
            '维系费': {
                'model_path': os.path.join(self.model_path, "附件5：维系费稽核清单.xlsx"),
                'output_path': os.path.join(self.output_path, "附件5：维系费稽核清单.xlsx"),
                'source_sheet_names': ['6-3维系费稽核问题清单模板'],
                'target_sheet_index': 1,
                'data_range': 'A:U',
                'start_row': 4,
                'file_keywords': ['维系费'],
                'text_columns': ['D']  # 维系费表需要文本格式的列
            }
        }

        self.cities = ['南宁', '柳州', '桂林', '玉林', '北海', '百色', '河池', '钦州', '贵港', '梧州', '防城港', '崇左', '来宾', '贺州']
        self.fees = ['电费', '场地费', '修理维护费', '发电费', '维系费']

    def find_city_file(self, city_name, keywords):
        """根据城市名和关键词查找文件"""
        for file_path in glob.glob(os.path.join(self.city_path, "**", "*.xls*"), recursive=True):
            file_name = os.path.basename(file_path)
            parent_dir = os.path.basename(os.path.dirname(file_path))

            if (city_name in parent_dir and
                    all(keyword in file_name for keyword in keywords)):
                print(f"  找到文件：{file_path}")
                return file_path
        return None

    def get_data_range(self, sheet, start_row, col_range):
        """获取数据范围，增加错误处理"""
        try:
            # 拆分列范围
            col_start, col_end = col_range.split(':')

            # 获取最后一行
            last_row = sheet.Cells(sheet.Rows.Count, col_start).End(-4162).Row  # -4162 表示xlUp

            # 安全检查
            if last_row < start_row:
                print(f"  警告：最后一行({last_row})小于起始行({start_row})")
                return None

            return sheet.Range(f'{col_start}{start_row}:{col_end}{last_row}')
        except Exception as e:
            print(f"  获取数据范围时出错: {str(e)}")
            return None

    def process_single_fee(self, xl, workbook_main, fee_type, city, source_sheet_name,
                           target_sheet_index, data_range, start_row, keywords, text_columns):
        """处理单个费用类型的数据，包含文本格式设置"""
        data_file = self.find_city_file(city, keywords)
        if not data_file:
            print(f"  警告：未找到{city}的{fee_type}文件")
            return 0

        workbook_data = None
        try:
            # 检查文件是否被锁定
            if self.is_file_locked(data_file):
                print(f"  错误：文件被锁定，请关闭后重试 - {data_file}")
                return 0

            # 打开数据文件（只读模式）
            workbook_data = xl.Workbooks.Open(
                Filename=data_file,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                UpdateLinks=0  # 不更新链接
            )

            # 通过名称获取源工作表
            try:
                sheet_data = workbook_data.Sheets(source_sheet_name)
                print(f"  已加载数据源工作表：{source_sheet_name}")
            except Exception as e:
                print(f"  错误：文件中找不到名为'{source_sheet_name}'的工作表 - {str(e)}")
                return 0

            # 关闭筛选
            if sheet_data.AutoFilterMode:
                try:
                    sheet_data.AutoFilterMode = False
                except Exception as e:
                    print(f"  关闭筛选时出错: {str(e)}")

            # 获取数据范围
            source_range = self.get_data_range(sheet_data, start_row, data_range)
            if not source_range:
                print(f"  警告：{city}的{fee_type}文件中没有有效数据")
                return 0

            data_rows = source_range.Rows.Count
            data_cols = source_range.Columns.Count
            print(f"  发现{city}的{fee_type}数据共{data_rows}行, {data_cols}列")

            # 获取主表对应的目标工作表
            try:
                sheet_main = workbook_main.Sheets(target_sheet_index)
                print(f"  已加载主表工作表（索引：{target_sheet_index}，名称：{sheet_main.Name}）")
            except Exception as e:
                print(f"  错误：主表中找不到索引为{target_sheet_index}的工作表 - {str(e)}")
                return 0

            # 解析列范围
            col_start, col_end = data_range.split(':')

            # 确定目标起始行
            try:
                target_start_row = sheet_main.Cells(sheet_main.Rows.Count, col_start).End(-4162).Row + 1
                if target_start_row < start_row:
                    target_start_row = start_row
            except Exception as e:
                print(f"  计算目标起始行时出错: {str(e)}")
                target_start_row = start_row

            # 计算数据结束行
            end_row = target_start_row + data_rows - 1

            # 直接读取值并写入，设置文本格式避免科学计数法
            try:
                values = source_range.Value
                if values is None:
                    print(f"  警告：{city}的{fee_type}没有可复制的数据")
                    return 0

                target_range = sheet_main.Range(
                    f'{col_start}{target_start_row}:{col_end}{end_row}'
                )

                # 针对需要文本格式的列进行设置
                for col in text_columns:
                    # 构建该列的完整范围
                    col_range = f"{col}{target_start_row}:{col}{end_row}"
                    # 设置为文本格式
                    sheet_main.Range(col_range).NumberFormat = '@'

                # 粘贴值
                target_range.Value = values
                return data_rows
            except Exception as e:
                print(f"  复制数据时出错: {str(e)}")
                # 尝试备选方案 - 使用剪贴板
                try:
                    source_range.Copy()
                    target_range = sheet_main.Range(f'{col_start}{target_start_row}')

                    # 先设置格式再粘贴
                    for col in text_columns:
                        col_range = f"{col}{target_start_row}:{col}{end_row}"
                        sheet_main.Range(col_range).NumberFormat = '@'

                    target_range.PasteSpecial(Paste=win32.constants.xlPasteValues)
                    xl.CutCopyMode = False
                    return data_rows
                except Exception as e2:
                    print(f"  备选复制方法也失败: {str(e2)}")
                    return 0

        except Exception as e:
            print(f"  处理{city}的{fee_type}时出错：{str(e)}")
            return 0
        finally:
            # 确保数据文件被关闭
            if workbook_data:
                try:
                    workbook_data.Close(SaveChanges=False)
                except Exception as e:
                    print(f"  关闭数据文件时出错：{str(e)}")
            # 释放资源
            time.sleep(0.5)

    def is_file_locked(self, file_path):
        """检查文件是否被锁定"""
        try:
            with open(file_path, 'a'):
                return False
        except IOError:
            return True

    def excel_process(self, fee_type):
        """处理指定费用类型的Excel文件"""
        config = self.fee_configs[fee_type]
        print(f'正在处理费用类型：{fee_type}')

        xl = None
        workbook_main = None
        try:
            pythoncom.CoInitialize()
            xl = win32.gencache.EnsureDispatch('Excel.Application')
            xl.Visible = False  # 可改为True进行调试
            xl.DisplayAlerts = False
            xl.AskToUpdateLinks = False
            xl.AutomationSecurity = 3  # 禁用宏安全提示

            if not os.path.exists(config['model_path']):
                print(f"错误：模板文件不存在 - {config['model_path']}")
                return

            # 打开主表文件
            workbook_main = xl.Workbooks.Open(
                config['model_path'],
                UpdateLinks=0,
                ReadOnly=False
            )

            # 处理修理维护费（多子表）
            if fee_type == '修理维护费':
                for subtype in config['subtypes']:
                    print(f"  开始处理子类型：{subtype['name']}（目标工作表索引：{subtype['target_sheet_index']}）")

                    for city in self.cities:
                        print(f"    处理城市：{city}")
                        self.process_single_fee(
                            xl, workbook_main,
                            subtype['name'], city,
                            subtype['source_sheet_name'][0],
                            subtype['target_sheet_index'],
                            subtype['data_range'],
                            subtype['start_row'],
                            subtype['file_keywords'],
                            subtype['text_columns']  # 传递文本列配置
                        )
            else:
                # 处理其他费用类型（单工作表）
                print(f"  开始处理（目标工作表索引：{config['target_sheet_index']}）")

                for city in self.cities:
                    print(f"    处理城市：{city}")
                    self.process_single_fee(
                        xl, workbook_main,
                        fee_type, city,
                        config['source_sheet_names'][0],
                        config['target_sheet_index'],
                        config['data_range'],
                        config['start_row'],
                        config['file_keywords'],
                        config['text_columns']  # 传递文本列配置
                    )

            # 保存并关闭主表
            if os.path.exists(config['output_path']):
                try:
                    os.remove(config['output_path'])
                except Exception as e:
                    print(f"  删除现有文件时出错：{str(e)}")
                    return

            workbook_main.SaveAs(config['output_path'])
            print(f'已保存汇总表：{config["output_path"]}')

        except Exception as e:
            print(f"处理{fee_type}时发生错误：{str(e)}")
        finally:
            if workbook_main:
                try:
                    workbook_main.Close(SaveChanges=False)
                except:
                    pass
            if xl:
                try:
                    xl.DisplayAlerts = True
                    xl.Quit()
                except:
                    pass
                del xl
            pythoncom.CoUninitialize()
            time.sleep(1)

    def main(self):
        for fee in self.fees:
            self.excel_process(fee)
            print(f"--- {fee} 处理完成 ---\n")
        print("所有费用类型处理完毕！")
        # 增加按回车键结束程序的功能
        input("请按回车键结束程序...")

if __name__ == "__main__":
    summarizer = FeeSummarizer()
    summarizer.main()
