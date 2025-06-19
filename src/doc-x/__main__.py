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
    from doc2docx import convert
except ImportError:
    missing_deps.append("doc2docx")

if missing_deps:
    rprint(f"[bold red]请先安装以下库: pip install {' '.join(missing_deps)}[/bold red]")
    sys.exit(1)

console = Console()

def get_doc_files(directory: str) -> List[Path]:
    """获取指定目录下的所有DOC文件"""
    path = Path(directory)
    return sorted(list(path.glob("*.doc")))  # 只处理.doc文件

def select_doc_files(files: List[Path]) -> List[Path]:
    """让用户选择要转换的DOC文件"""
    if not files:
        console.print("[bold red]当前目录下没有找到DOC文件[/bold red]")
        return []
    
    table = Table(title="可用的DOC文件")
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
            "[yellow]请输入要转换的DOC文件序号(用逗号分隔多个序号，如1,3,5)，输入0全选[/yellow]"
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

def convert_doc_to_docx(doc_path: Path, output_folder: Path) -> Path:
    """将DOC转换为DOCX"""
    # 使用绝对路径
    doc_path = doc_path.resolve()
    output_folder = output_folder.resolve()
    
    # DOCX输出路径
    docx_path = output_folder / f"{doc_path.stem}.docx"
    
    try:
        # 使用doc2docx库进行转换
        convert(str(doc_path), str(docx_path))
        return docx_path
    except Exception as e:
        console.print(f"[bold red]转换 {doc_path.name} 为DOCX时出错: {str(e)}[/bold red]")
        raise

def main():
    console.print(Panel.fit("[bold blue]DOC文件转DOCX工具[/bold blue]", border_style="green"))
    
    # 获取当前目录
    current_dir = Prompt.ask("[yellow]请输入DOC文件所在目录[/yellow]", default=str(Path.cwd()))
    
    # 获取DOC文件列表
    doc_files = get_doc_files(current_dir)
    
    if not doc_files:
        console.print(f"[bold red]在 {current_dir} 目录下未找到DOC文件[/bold red]")
        return
    
    # 用户选择文件
    selected_files = select_doc_files(doc_files)
    
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
    output_dir = Path(Prompt.ask("[yellow]请输入DOCX输出目录[/yellow]", default=str(Path(current_dir) / "DOCX输出")))
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 转换DOC为DOCX
    file_count = len(selected_files)
    success_count = 0
    failed_files = []
    
    with Progress(
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
    ) as progress:
        convert_task = progress.add_task("[green]转换DOC为DOCX...", total=file_count)
        
        for i, doc_file in enumerate(selected_files, 1):
            progress.update(convert_task, description=f"[green]转换文件 {i}/{file_count}: {doc_file.name}")
            try:
                docx_file = convert_doc_to_docx(doc_file, output_dir)
                success_count += 1
                console.print(f"[green]✓ {doc_file.name} -> {docx_file.name}[/green]")
            except Exception as e:
                failed_files.append((doc_file.name, str(e)))
                console.print(f"[red]✗ {doc_file.name} 转换失败[/red]")
            progress.update(convert_task, advance=1)
    
    # 显示转换结果
    console.print(f"\n[bold green]转换完成![/bold green]")
    console.print(f"[green]成功转换: {success_count}/{file_count} 个文件[/green]")
    
    if failed_files:
        console.print(f"[red]失败文件: {len(failed_files)} 个[/red]")
        for filename, error in failed_files:
            console.print(f"  [red]- {filename}: {error}[/red]")
    
    console.print(f"[cyan]DOCX文件已保存到: {output_dir}[/cyan]")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print("[yellow]\n操作已取消[/yellow]")
    except Exception as e:
        console.print(f"[bold red]发生错误: {str(e)}[/bold red]")
        raise