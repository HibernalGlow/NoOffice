import re
import os
from pathlib import Path
from rich.console import Console
from rich.prompt import Prompt, Confirm
from rich.table import Table
from rich.panel import Panel

console = Console()

def convert_to_markdown_table(text):
    """将交易流水文本转换为Markdown表格"""
    
    # 定义表头
    headers = [
        "交易号", "商户订单号", "交易创建时间", "付款时间", "最近修改时间", 
        "交易来源地", "类型", "用户信息", "交易对方", "商品信息", 
        "消费名称", "金额（元）", "支付方式", "收/支", "交易状态", "备注"
    ]
    
    # 创建表头
    markdown_table = "| " + " | ".join(headers) + " |\n"
    markdown_table += "|" + "---|" * len(headers) + "\n"
    
    lines = text.split('\n')
    for line in lines:
        if '18868728387' in line:
            # 简化的数据提取（需要根据实际格式调整）
            parts = line.split()
            if len(parts) >= 16:
                row_data = parts[:16]  # 取前16列
                markdown_table += "| " + " | ".join(row_data) + " |\n"
    
    return markdown_table

def get_md_files(directory):
    """获取指定目录下的所有Markdown文件"""
    path = Path(directory)
    return sorted(list(path.glob("*.md")))

def select_md_file(files):
    """让用户选择要处理的Markdown文件"""
    if not files:
        console.print("[bold red]当前目录下没有找到Markdown文件[/bold red]")
        return None
    
    table = Table(title="可用的Markdown文件")
    table.add_column("序号", justify="right", style="cyan")
    table.add_column("文件名", style="green")
    table.add_column("文件大小", justify="right")
    
    for i, file in enumerate(files, 1):
        size = file.stat().st_size / 1024  # KB
        size_str = f"{size:.2f} KB"
        table.add_row(str(i), file.name, size_str)
    
    console.print(table)
    
    while True:
        selection = Prompt.ask("[yellow]请输入要处理的Markdown文件序号[/yellow]")
        
        try:
            index = int(selection) - 1
            if 0 <= index < len(files):
                return files[index]
            else:
                console.print("[bold red]序号超出范围，请重新输入[/bold red]")
        except ValueError:
            console.print("[bold red]输入格式错误，请输入数字[/bold red]")

def main():
    console.print(Panel.fit("[bold blue]交易流水表格转换工具[/bold blue]", border_style="green"))
    
    # 获取当前目录
    current_dir = Path(__file__).parent
    console.print(f"[cyan]当前目录: {current_dir}[/cyan]")
    
    # 获取Markdown文件列表
    md_files = get_md_files(current_dir)
    
    if not md_files:
        console.print(f"[bold red]在 {current_dir} 目录下未找到Markdown文件[/bold red]")
        return
    
    # 用户选择文件
    selected_file = select_md_file(md_files)
    
    if not selected_file:
        return
    
    console.print(f"[green]已选择文件: {selected_file.name}[/green]")
    
    # 确认操作
    if not Confirm.ask("[yellow]确认处理此文件?[/yellow]"):
        console.print("[yellow]操作已取消[/yellow]")
        return
    
    try:
        # 读取文件内容
        with open(selected_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 转换为表格
        table_content = convert_to_markdown_table(content)
        
        # 输出文件路径
        output_file = selected_file.parent / f"{selected_file.stem}_table.md"
        
        # 写入转换后的表格
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(table_content)
        
        console.print(f"[bold green]转换完成! 表格已保存到: {output_file}[/bold green]")
        
    except Exception as e:
        console.print(f"[bold red]处理文件时出错: {str(e)}[/bold red]")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print("[yellow]\n操作已取消[/yellow]")
    except Exception as e:
        console.print(f"[bold red]发生错误: {str(e)}[/bold red]")