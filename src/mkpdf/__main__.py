import os
import subprocess
from pathlib import Path
import rich
from rich.prompt import Prompt, Confirm
from rich.console import Console
from rich.panel import Panel
from rich.columns import Columns

def convert_file_to_md(file_path, output_dir=None, file_type=None):
    """将指定格式文件转换为Markdown"""
    if output_dir is None:
        output_dir = os.path.dirname(file_path)
    
    base_name = os.path.basename(file_path)
    file_name = os.path.splitext(base_name)[0]
    output_path = os.path.join(output_dir, f"{file_name}.md")
    
    print(f"正在转换: {file_path} -> {output_path}")
    
    try:
        # 使用markitdown命令行工具进行转换
        subprocess.run(['markitdown', file_path, '-o', output_path], check=True)
        print(f"成功转换: {output_path}")
        return True
    except Exception as e:
        print(f"转换失败: {file_path}, 错误: {str(e)}")
        return False

def main():
    console = Console()
    
    # 使用rich交互获取文件夹路径
    console.print(Panel("[bold blue]文档转Markdown工具[/bold blue]", 
                        subtitle="支持多种文档格式转换"))
    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = Prompt.ask("请输入文档文件所在路径", default=current_dir)
    
    # 选择文件格式
    format_choices = {
        "1": {"name": "PDF文件 (.pdf)", "ext": [".pdf"]},
        "2": {"name": "Word文档 (.doc, .docx)", "ext": [".doc", ".docx"]},
        "3": {"name": "PowerPoint演示文稿 (.ppt, .pptx)", "ext": [".ppt", ".pptx"]},
        "4": {"name": "手动输入扩展名", "ext": []}
    }
    
    # 显示格式选项
    console.print("\n[bold cyan]请选择要转换的文件格式:[/bold cyan]")
    for key, value in format_choices.items():
        console.print(f"  {key}. {value['name']}")
    
    # 获取用户选择
    choice = Prompt.ask("请输入选项编号", choices=["1", "2", "3", "4"], default="1")
    selected_format = format_choices[choice]
    
    # 如果选择手动输入
    if choice == "4":
        extensions_input = Prompt.ask("请输入文件扩展名 (用逗号分隔，如: .pdf,.docx)")
        extensions = [ext.strip() for ext in extensions_input.split(",")]
        selected_format["ext"] = extensions
    
    # 创建output文件夹
    output_dir = Prompt.ask("请输入输出路径", default=os.path.join(input_dir, "markdown_output"))
    os.makedirs(output_dir, exist_ok=True)
    
    # 获取所有符合选定扩展名的文件
    extensions = [ext.lower() for ext in selected_format["ext"]]
    target_files = []
    
    for file in os.listdir(input_dir):
        file_lower = file.lower()
        for ext in extensions:
            if file_lower.endswith(ext):
                target_files.append(file)
                break
    
    if not target_files:
        console.print(f"[bold red]指定文件夹中没有找到{selected_format['name']}![/bold red]")
        return
    
    console.print(f"[green]找到 {len(target_files)} 个{selected_format['name']}文件[/green]")
    
    # 确认是否开始转换
    if Confirm.ask("是否开始转换?", default=True):
        # 转换所有文件
        with console.status(f"[bold green]正在转换{selected_format['name']}文件...") as status:
            success_count = 0
            for file in target_files:
                file_path = os.path.join(input_dir, file)
                if convert_file_to_md(file_path, output_dir):
                    success_count += 1
        
        console.print(f"[bold green]转换完成! {success_count}/{len(target_files)} 个文件成功转换[/bold green]")
    else:
        console.print("[yellow]已取消转换[/yellow]")

if __name__ == "__main__":
    main()