# -*- coding: utf-8 -*-
import os, sys, glob
sys.path.insert(0, r"C:\Users\admin\Desktop\报告生成器\app")
import report_engine as E

GT = r"C:\Users\admin\Desktop\测试项目\O1D\高温工作"
ZD = r"C:\Users\admin\Desktop\测试项目\O1D\振动"

def imgs(folder, n=None):
    fs = sorted(glob.glob(os.path.join(folder, "*.jpg")))
    if n: fs = fs[:n]
    return [{"path": p, "caption": ""} for p in fs]

project = {
  "info": {
    "report_no": "YJ/SYBG-20260527001",
    "sample_name": "O1D大通DMS一体机", "sample_no": "C00771959",
    "sample_model": "65左舵", "verify_phase": "其他",
    "client_name": "深圳佑驾创新科技有限公司",
    "client_addr": "深圳市福田区沙头街道上沙社区滨河大道9285号中......",
    "maker_name": "深圳佑驾创新科技有限公司",
    "maker_addr": "深圳市福田区沙头街道上沙社区滨河大道9285号中......",
    "sample_qty": "2", "sample_way": "客户送样",
    "recv_date": "2026.02.09", "commission_no": "ME/WTD20251230012",
    "test_date_range": "2026.02.28-2026.05.27",
    "lab_name": "深圳佑驾创新科技股份有限公司实验中心",
    "test_items": "参考客户要求", "test_basis": "参考客户要求", "remark": "",
  },
  "tests": [
    {
      "title": "高温耐久测试", "standard": "参考客户要求", "sample_no": "1#-2#",
      "start_date": "2026.02.28", "end_date": "2026.05.25", "overall_result": "合格",
      "sample_name": "O1D大通DMS一体机",
      "equipment": [
        {"name":"恒温恒湿箱","model":"XB-OTS-150B-C","mgmt_no":"ME/LAB-003","cal_valid":"2025.11.17-2026.11.16"},
        {"name":"直流稳压电源","model":"IT6722A","mgmt_no":"ME/LAB-029","cal_valid":"2025.11.17-2026.11.16"},
      ],
      "env":"18℃-28℃、25%RH-75%RH","test_date":"2026.02.28-2026.05.25",
      "condition":"测试温度:90℃/测试时间:2021h/操作模式","requirement":"功能等级A。",
      "samples":[
        {"no":"1#","result":"试验前，样品功能正常，报文正常滚动；试验后功能正常。","conclusion":"合格"},
        {"no":"2#","result":"试验前，样品功能正常，报文正常滚动；试验后功能正常。","conclusion":"合格"},
      ],
      "image_groups":[
        {"title":"试验前图片","images":imgs(GT,2)},
        {"title":"试验中图片","images":imgs(GT,3)},
        {"title":"试验后图片","images":imgs(GT,1)},
      ],
    },
    {
      "title": "振动测试", "standard": "参考客户要求", "sample_no": "1#-2#",
      "start_date": "2026.05.25", "end_date": "2026.05.27", "overall_result": "合格",
      "sample_name": "O1D大通DMS一体机",
      "equipment": [
        {"name":"振动试验台","model":"DC-1000-15","mgmt_no":"ME/LAB-015","cal_valid":"2025.11.17-2026.11.16"},
        {"name":"直流稳压电源","model":"IT6722A","mgmt_no":"ME/LAB-029","cal_valid":"2025.11.17-2026.11.16"},
        {"name":"三综合试验箱","model":"XB-OTS-408F-B","mgmt_no":"ME/LAB-014","cal_valid":"2025.11.17-2026.11.16"},
      ],
      "env":"18℃-28℃、25%RH-75%RH","test_date":"2026.05.25-2026.05.27",
      "condition":"随机振动 X/Y/Z 三轴","requirement":"工作模式3.2：功能等级 A",
      "samples":[
        {"no":"1#","result":"X/Y/Z轴振动试验前后，样品外观正常，报文正常滚动。","conclusion":"合格"},
        {"no":"2#","result":"X/Y/Z轴振动试验前后，样品外观正常，报文正常滚动。","conclusion":"合格"},
      ],
      "image_groups":[
        {"title":"X轴图片","images":imgs(os.path.join(ZD,"x"))},
        {"title":"Y轴图片","images":imgs(os.path.join(ZD,"y"))},
        {"title":"Z轴图片","images":imgs(os.path.join(ZD,"z"))},
      ],
    },
  ],
}

out = r"C:\Users\admin\Desktop\报告生成器\输出\_测试输出.docx"
E.generate(project, out)
print("GENERATED", os.path.getsize(out))
