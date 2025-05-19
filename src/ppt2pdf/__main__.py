import os
import sys
from pathlib import Path
from typing import List, Optional
import tempfile
import shutil

from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt, Confirm
from rich.table import Table
from rich.progress import Progress, TextColumn, BarColumn, TaskProgressColumn
from rich import print as rprint

# 检查必要的依赖
missing_deps = []
try:
    from pptx import Presentation
except ImportError:
    missing_deps.append("python-pptx")

try:
    import comtypes.client
except ImportError:
    missing_deps.append("comtypes")

if missing_deps:
    rprint(f"[bold red]请先安装以下库: pip install {' '.join(missing_deps)}[/bold red]")
    sys.exit(1)

console = Console()

def get_ppt_files(directory: str) -> List[Path]:
    """获取指定目录下的所有PPT文件"""
    path = Path(directory)
    return sorted(list(path.glob("*.pptx")))  # 只处理pptx文件

def select_ppt_files(files: List[Path]) -> List[Path]:
    """让用户选择要转换的PPT文件"""
    if not files:
        console.print("[bold red]当前目录下没有找到PPTX文件[/bold red]")
        return []
    
    table = Table(title="可用的PPTX文件")
    table.add_column("序号", justify="right", style="cyan")
    table.add_column("文件名", style="green")
    table.add_column("文件大小", justify="right")
    
    for i, file in enumerate(files, 1):
        size = file.stat().st_size / 1024  # KB
        size_str = f"{size:.2f} KB"
        table.add_row(str(i), file.name, size_str)
    
    console.print(table)
    
    while True:
        selection = Prompt.ask(
            "[yellow]请输入要转换的PPTX文件序号(用逗号分隔多个序号，如1,3,5)，输入0全选[/yellow]"
        )
        
        if selection.lower() == "0":
            return files
        
        try:
            indices = [int(idx.strip()) - 1 for idx in selection.split(",")]
            selected_files = [files[idx] for idx in indices if 0 <= idx < len(files)]
            if not selected_files:
                console.print("[bold red]未选择任何有效文件，请重新选择[/bold red]")
                continue
            
            return selected_files
        except (ValueError, IndexError):
            console.print("[bold red]输入格式错误，请重新输入[/bold red]")

def convert_pptx_to_pdf(pptx_path: Path, output_folder: Path) -> Path:
    """将PPTX转换为PDF"""
    # 使用绝对路径
    pptx_path = pptx_path.resolve()
    output_folder = output_folder.resolve()
    
    # PDF输出路径
    pdf_path = output_folder / f"{pptx_path.stem}.pdf"
    
    powerpoint = None
    
    try:
        # 使用COM对象进行转换 - 静默模式
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        # 先创建对象，然后设置属性
        try:
            powerpoint.Visible = 0  # 设置为0使PowerPoint在后台运行
        except:
            # 如果无法设置可见性，继续尝试转换
            console.print("[yellow]警告: 无法将PowerPoint设置为后台模式[/yellow]")
            
        # 打开演示文稿
        try:
            presentation = powerpoint.Presentations.Open(str(pptx_path), WithWindow=False)
        except:
            # 如果WithWindow失败，尝试不带参数打开
            presentation = powerpoint.Presentations.Open(str(pptx_path))
            
        # PDF格式代码为32
        presentation.SaveAs(str(pdf_path), 32)
        presentation.Close()
        return pdf_path
    except Exception as e:
        console.print(f"[bold red]转换 {pptx_path.name} 为PDF时出错: {str(e)}[/bold red]")
        raise
    finally:
        # 安全地关闭PowerPoint
        if powerpoint:
            try:
                powerpoint.Quit()
            except:
                pass

def main():
    console.print(Panel.fit("[bold blue]PPT文件转PDF工具[/bold blue]", border_style="green"))
    
    # 获取当前目录
    current_dir = Prompt.ask("[yellow]请输入PPTX文件所在目录[/yellow]", default=str(Path.cwd()))
    
    # 获取PPT文件列表
    ppt_files = get_ppt_files(current_dir)
    
    if not ppt_files:
        console.print(f"[bold red]在 {current_dir} 目录下未找到PPTX文件[/bold red]")
        return
    
    # 用户选择文件
    selected_files = select_ppt_files(ppt_files)
    
    if not selected_files:
        return
    
    # 显示已选择的文件
    console.print("[bold green]已选择以下文件进行处理:[/bold green]")
    for i, file in enumerate(selected_files, 1):
        console.print(f"  {i}. [cyan]{file.name}[/cyan]")
    
    # 确认操作
    if not Confirm.ask("[yellow]确认处理这些文件?[/yellow]"):
        console.print("[yellow]操作已取消[/yellow]")
        return
    
    # 询问输出目录
    output_dir = Path(Prompt.ask("[yellow]请输入PDF输出目录[/yellow]", default=str(Path(current_dir) / "PDF输出")))
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 转换PPTX为PDF
    file_count = len(selected_files)
    
    with Progress(
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
    ) as progress:
        convert_task = progress.add_task("[green]转换PPTX为PDF...", total=file_count)
        
        for i, pptx_file in enumerate(selected_files, 1):
            progress.update(convert_task, description=f"[green]转换文件 {i}/{file_count}: {pptx_file.name}")
            pdf_file = convert_pptx_to_pdf(pptx_file, output_dir)
            progress.update(convert_task, advance=1)
    
    console.print(f"[bold green]转换完成! PDF文件已保存到: [/bold green][cyan]{output_dir}[/cyan]")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print("[yellow]\n操作已取消[/yellow]")
    except Exception as e:
        console.print(f"[bold red]发生错误: {str(e)}[/bold red]")
        raise