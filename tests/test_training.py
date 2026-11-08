"""作者 Wangyuan，2026-10-29。真实生成工作簿、回读结果及图像的离线回归。"""
import json,subprocess,sys,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
import pandas as pd
from PIL import Image,ImageDraw,ImageStat
from openpyxl import load_workbook
from training import *
from demo import synthetic_frame,example_config,run

class TrainingTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name);self.input=self.root/"input.xlsx";self.frame=synthetic_frame();self.config=example_config()
    def tearDown(self):self.tmp.cleanup()
    def clean(self,frame=None,**kwargs):
        (self.frame if frame is None else frame).to_excel(self.input,index=False)
        return clean_ledger(self.input,**kwargs)
    def result(self):
        frame,notes=self.clean();return analyze(frame,self.config,notes)
    def test_actual_totals_and_weighted_score(self):
        result=self.result();self.assertEqual(result.kpis["total_hours"],6);self.assertEqual(result.kpis["total_cost"],300);self.assertEqual(result.kpis["participants"],35);self.assertAlmostEqual(result.kpis["weighted_score"],80)
        self.assertEqual(result.trend.hours.tolist(),[2,4]);self.assertEqual(result.excellent.instructor.tolist(),["Instructor-A"]);self.assertEqual(result.excellent.total_hours.iloc[0],3)
    def test_duplicate_excel_headers(self):
        self.clean();wb=load_workbook(self.input);wb.active.cell(1,2,"course");wb.active.cell(1,3,"course");wb.save(self.input);wb.close()
        with self.assertRaises(ValueError):clean_ledger(self.input)
    def test_empty_end_to_end_has_no_poster(self):
        self.frame=self.frame.iloc[:0];self.clean();reply=run(self.input,self.root/"empty",self.config);self.assertFalse(reply["poster"]);self.assertTrue((self.root/"empty"/"dashboard.png").exists());self.assertFalse((self.root/"empty"/"selected-poster.png").exists())
    def test_config_type_rejected(self):
        self.clean()
        with self.assertRaises(ValueError):run(self.input,self.root/"bad",[])
        self.assertFalse((self.root/"bad").exists())
    def test_explicit_aliases(self):
        self.frame=self.frame.rename(columns={"course":"Sample Course"});f,_=self.clean(aliases={"Sample Course":"course"});self.assertIn("course",f)
    def test_missing_columns(self):
        self.frame=self.frame.drop(columns="hours")
        with self.assertRaises(ValueError):self.clean()
    def test_duplicate_columns(self):
        with self.assertRaises(ValueError):rename_columns(pd.DataFrame([[1,2]],columns=["course"," course "]))
        with self.assertRaises(ValueError):rename_columns(pd.DataFrame([[1,2]],columns=["a","b"]),{"a":"course","b":"course"})
    def test_summary_only_explicit(self):
        self.frame.loc[0,"course"]="TOTAL"
        f,_=self.clean();self.assertEqual(len(f),3)
        f,n=self.clean(summary_labels=("TOTAL",));self.assertEqual(len(f),2);self.assertTrue(n)
    def test_numeric_units(self):
        self.frame=self.frame.astype({"fee":object,"hours":object})
        self.frame.loc[0,"fee"]="1,000元";self.frame.loc[0,"hours"]="2小时"
        f,_=self.clean();self.assertEqual(f.fee_num.iloc[0],1000);self.assertEqual(f.hours_num.iloc[0],2)
    def test_bad_numbers_rejected(self):
        for value in ["bad",float("inf"),-1]:
            self.frame=synthetic_frame().astype({"hours":object});self.frame.loc[0,"hours"]=value
            with self.assertRaises(ValueError):self.clean()
    def test_missing_numbers_explicit_policy(self):
        self.frame.loc[0,"hours"]=None
        with self.assertRaises(ValueError):self.clean()
        f,n=self.clean(missing_numbers="zero");self.assertEqual(f.hours_num.iloc[0],0);self.assertTrue(n)
    def test_fractional_people_rejected(self):
        self.frame=self.frame.astype({"participants":float})
        self.frame.loc[0,"participants"]=1.5
        with self.assertRaises(ValueError):self.clean()
    def test_missing_score_is_unknown(self):
        self.frame["satisfaction"]=None;r=self.result();self.assertIsNone(r.kpis["weighted_score"]);self.assertTrue(r.excellent.empty);self.assertEqual(r.quality_bins["count"].sum(),0)
        save_summary(r,self.root/"summary.json");self.assertIsNone(json.loads((self.root/"summary.json").read_text())["kpis"]["weighted_score"])
    def test_zero_weights_no_unweighted_fallback(self):
        self.frame["participants"]=0;self.assertIsNone(self.result().kpis["weighted_score"])
    def test_invalid_dates_are_not_coerced(self):
        for value in ["2030-99-01","2030-02-30","2030-00","garbage","02月"]:self.assertTrue(pd.isna(parse_month_value(value)))
        self.assertEqual(parse_month_value("2032-02-29"),pd.Timestamp("2032-02-01"))
        self.assertEqual(parse_month_value("2030年2月"),pd.Timestamp("2030-02-01"))
        self.assertEqual(parse_month_value("2月",2031),pd.Timestamp("2031-02-01"))
    def test_year_not_inferred(self):
        self.frame.loc[0,"date"]="1月";f,n=self.clean();self.assertTrue(pd.isna(f.month.iloc[0]));self.assertTrue(n)
        f,_=self.clean(default_year=2035);self.assertEqual(f.month.iloc[0].year,2035)
    def test_cross_year_labels(self):
        self.frame.loc[0,"date"]="2029-02-01";r=self.result();self.assertEqual(r.trend.label.tolist(),["2029-02","2030-02"])
    def test_all_unknown_dates_no_selection(self):
        self.frame["date"]="unknown";r=self.result();self.assertTrue(r.trend.empty);self.assertTrue(r.excellent.empty)
    def test_explicit_selection_month_and_threshold(self):
        self.config["latest_month"]="2030-01";r=self.result();self.assertEqual(r.excellent.total_hours.iloc[0],2)
        self.config["selection_rule"]["min_satisfaction"]=99;self.assertTrue(self.result().excellent.empty)
    def test_score_bin_edges_and_invalid_rules(self):
        self.frame["satisfaction"]=[0,50,100];r=self.result();self.assertEqual(r.quality_bins["count"].tolist(),[2,0,1])
        for edges in [[0,100,50],[0,float("inf")],[0,40]]:
            self.config["quality_edges"]=edges
            with self.assertRaises(ValueError):self.result()
        self.config=example_config();self.config["selection_rule"]["top_n"]=0
        with self.assertRaises(ValueError):self.result()
    def test_empty_workbook_with_schema(self):
        self.frame=self.frame.iloc[:0];r=self.result();self.assertEqual(r.kpis["participants"],0);self.assertIsNone(r.kpis["weighted_score"]);self.assertTrue(r.excellent.empty)
    def test_missing_identifier_rejected(self):
        self.frame.loc[0,"instructor"]=None
        with self.assertRaises(ValueError):self.clean()
    def test_empty_selection_no_fabricated_poster(self):
        self.config["selection_rule"]["min_satisfaction"]=99;r=self.result()
        with self.assertRaises(ValueError):render_poster(r,self.config,self.root/"none.png")
        self.assertFalse((self.root/"none.png").exists())
    def test_export_formula_text(self):
        self.frame.loc[0,"instructor"]="@sample";self.frame.loc[1,"instructor"]="@sample";r=self.result();save_analysis_workbook(r,self.root/"analysis.xlsx");save_excellent_csv(r,self.root/"selected.csv")
        wb=load_workbook(self.root/"analysis.xlsx")
        try:
            cells=[c for row in wb["selected"] for c in row];self.assertFalse(any(c.data_type=="f" for c in cells));self.assertTrue(any(c.value=="'@sample" for c in cells))
        finally:wb.close()
        self.assertIn("'@sample",(self.root/"selected.csv").read_text())
    def test_pie_other_keeps_denominator(self):
        form=pd.DataFrame({"form":list("abcdef"),"share_value":[1]*6});draw=ImageDraw.Draw(Image.new("RGB",(900,600)))
        with patch("training.draw_text") as captured:draw_pie(draw,(0,0,800,500),form)
        labels=[call.args[2] for call in captured.call_args_list];self.assertTrue(any("other 33.3%" in str(x) for x in labels))
    def test_end_to_end_workbook_images_and_json(self):
        self.clean();out=self.root/"results";reply=run(self.input,out,self.config);self.assertTrue(reply["poster"])
        data=json.loads((out/"summary.json").read_text());self.assertEqual(data["kpis"]["participants"],35)
        wb=load_workbook(out/"analysis.xlsx",read_only=True)
        try:self.assertEqual(wb["monthly"].max_row,3);self.assertIn("score_bins",wb.sheetnames)
        finally:wb.close()
        for name,size in [("dashboard.png",(2400,1600)),("selected-poster.png",(1080,1620))]:
            with Image.open(out/name) as image:self.assertEqual(image.size,size);self.assertGreater(sum(ImageStat.Stat(image).stddev),10)
            with Image.open(out/name) as image:image.verify()
        with self.assertRaises(ValueError):run(self.input,out,self.config)
    def test_cli_synthetic_and_no_overwrite(self):
        dest=self.root/"demo";cmd=[sys.executable,"-B","demo.py","--demo","--output",str(dest)]
        proc=subprocess.run(cmd,text=True,capture_output=True);self.assertEqual(proc.returncode,0,proc.stderr);self.assertTrue(json.loads(proc.stdout)["poster"])
        self.assertEqual(subprocess.run(cmd,capture_output=True).returncode,2)

if __name__=="__main__":unittest.main()
