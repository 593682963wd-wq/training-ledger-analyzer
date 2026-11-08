"""作者 Wangyuan，2026-10-29。只处理明确文件的离线培训分析入口。"""
from __future__ import annotations
import argparse,json
from pathlib import Path
import pandas as pd
from training import clean_ledger,analyze,font_context,render_dashboard,render_poster,save_summary,save_excellent_csv,save_analysis_workbook

def synthetic_frame()->pd.DataFrame:
    """新造样例，不含来源人员、课程或组织。"""
    return pd.DataFrame([
        {"report_dept":"Example-A","course":"Synthetic geometry","date":"2030-01-12","instructor":"Instructor-A","participants":10,"hours":2,"satisfaction":70,"fee":100,"form":"Workshop"},
        {"report_dept":"Example-A","course":"Synthetic colors","date":"2030-02-13","instructor":"Instructor-A","participants":20,"hours":3,"satisfaction":90,"fee":150,"form":"Workshop"},
        {"report_dept":"Example-B","course":"Synthetic symbols","date":"2030-02-14","instructor":"Instructor-B","participants":5,"hours":1,"satisfaction":60,"fee":50,"form":"Practice"},
    ])

def example_config()->dict:
    """只用于合成示例的任意量表/阈值，无业务或监管含义。"""
    return {"selection_rule":{"min_satisfaction":65,"top_n":3},"quality_edges":[0,50,80,100],"poster":{"title_template":"Selected: {course}","slogan":"Synthetic demonstration only","footer":"Synthetic data"}}

def run(input_path:Path,output:Path,config:dict,*,font_path:Path|None=None,sheet:str|None=None)->dict:
    """先验证与分析再新建输出目录；不覆盖已有结果。"""
    if not isinstance(config,dict):raise ValueError("配置必须为对象")
    if output.exists():raise ValueError("输出目录已存在，请使用新目录")
    frame,notes=clean_ledger(input_path,sheet,aliases=config.get("columns"),default_year=config.get("default_year"),summary_labels=tuple(config.get("summary_labels",[])),missing_numbers=config.get("missing_numbers","error"))
    result=analyze(frame,config,notes)
    # 字体错误在创建输出目录前报告。
    with font_context(font_path):
        output.mkdir(parents=True,exist_ok=False)
        render_dashboard(result,output/"dashboard.png")
        if not result.excellent.empty:render_poster(result,config,output/"selected-poster.png")
        else:result.notes.append("No eligible candidates: poster intentionally omitted")
        save_summary(result,output/"summary.json")
        save_excellent_csv(result,output/"selected.csv")
        save_analysis_workbook(result,output/"analysis.xlsx")
    return {"output":str(output),"rows":len(frame),"selected":len(result.excellent),"poster":not result.excellent.empty,"notes":result.notes}

def main()->None:
    parser=argparse.ArgumentParser(description="Offline training analytics, explicit inputs only")
    parser.add_argument("--demo",action="store_true",help="create only synthetic inputs")
    parser.add_argument("--input",type=Path);parser.add_argument("--config",type=Path);parser.add_argument("--output",type=Path,required=True);parser.add_argument("--font",type=Path);parser.add_argument("--sheet")
    args=parser.parse_args()
    try:
        if args.demo:
            if args.input or args.config:raise ValueError("--demo不能同时指定输入和配置")
            if args.output.exists():raise ValueError("输出目录已存在")
            args.output.mkdir(parents=True);input_path=args.output/"synthetic.xlsx";synthetic_frame().to_excel(input_path,index=False)
            result=run(input_path,args.output/"results",example_config(),font_path=args.font)
        else:
            if not args.input or not args.config:raise ValueError("需要--input和--config，或使用--demo")
            config=json.loads(args.config.read_text());result=run(args.input,args.output,config,font_path=args.font,sheet=args.sheet)
        print(json.dumps(result,ensure_ascii=False,allow_nan=False))
    except (ValueError,TypeError,KeyError,OSError) as exc:parser.exit(2,f"error: {exc}\n")
if __name__=="__main__":main()
