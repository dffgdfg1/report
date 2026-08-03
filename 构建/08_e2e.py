# -*- coding: utf-8 -*-
"""端到端：用 urllib 走 保存->上传->生成 全流程（模拟前端）。"""
import os, json, glob, urllib.request, uuid, mimetypes

BASE="http://127.0.0.1:8731"
NAME="E2E测试_O1D"
GT=r"C:\Users\admin\Desktop\测试项目\O1D\高温工作"
ZD=r"C:\Users\admin\Desktop\测试项目\O1D\振动"

def post_json(path, obj):
    data=json.dumps(obj).encode("utf-8")
    req=urllib.request.Request(BASE+path, data=data, headers={"Content-Type":"application/json"})
    return json.loads(urllib.request.urlopen(req).read())

def get(path):
    return json.loads(urllib.request.urlopen(BASE+path).read())

def upload(files):
    boundary="----e2e"+uuid.uuid4().hex
    body=b""
    def add_field(name,val):
        nonlocal body
        body+=("--"+boundary+"\r\n").encode()
        body+=('Content-Disposition: form-data; name="%s"\r\n\r\n'%name).encode()
        body+=(val+"\r\n").encode()
    add_field("project",NAME)
    for fp in files:
        fn=os.path.basename(fp)
        body+=("--"+boundary+"\r\n").encode()
        body+=('Content-Disposition: form-data; name="files"; filename="%s"\r\n'%fn).encode()
        body+=("Content-Type: image/jpeg\r\n\r\n").encode()
        body+=open(fp,"rb").read()+b"\r\n"
    body+=("--"+boundary+"--\r\n").encode()
    req=urllib.request.Request(BASE+"/api/upload", data=body,
        headers={"Content-Type":"multipart/form-data; boundary="+boundary})
    return json.loads(urllib.request.urlopen(req).read())

# 1) 建项目（空图）
info={"report_no":"YJ/SYBG-20260527001","sample_name":"O1D大通DMS一体机","sample_no":"C00771959",
 "sample_model":"65左舵","verify_phase":"其他","client_name":"深圳佑驾创新科技有限公司",
 "sample_qty":"2","sample_way":"客户送样","recv_date":"2026.02.09",
 "commission_no":"ME/WTD20251230012","test_date_range":"2026.02.28-2026.05.27",
 "lab_name":"深圳佑驾创新科技股份有限公司实验中心","test_items":"参考客户要求","test_basis":"参考客户要求"}
proj={"name":NAME,"info":info,"tests":[]}
print("save1", post_json("/api/save", proj))

# 2) 上传高温图片
gt_files=sorted(glob.glob(os.path.join(GT,"*.jpg")))
up1=upload(gt_files)
print("upload 高温", up1["ok"], len(up1["files"]))
x_files=sorted(glob.glob(os.path.join(ZD,"x","*.jpg")))
up2=upload(x_files)
print("upload 振动X", up2["ok"], len(up2["files"]))

def to_imgs(u): return [{"file":f["file"],"orig":f["orig"],"caption":""} for f in u["files"]]

t1={"title":"高温耐久测试","sample_no":"1#-2#","standard":"参考客户要求","start_date":"2026.02.28",
 "end_date":"2026.05.25","overall_result":"合格","sample_name":"O1D大通DMS一体机",
 "env":"18℃-28℃、25%RH-75%RH","test_date":"2026.02.28-2026.05.25","condition":"测试温度:90℃/2021h","requirement":"功能等级A。",
 "equipment":[{"name":"恒温恒湿箱","model":"XB-OTS-150B-C","mgmt_no":"ME/LAB-003","cal_valid":"2025.11.17-2026.11.16"}],
 "samples":[{"no":"1#","result":"试验前后功能正常","conclusion":"合格"}],
 "image_groups":[{"title":"试验前图片","images":to_imgs(up1)[:2]},{"title":"试验中图片","images":to_imgs(up1)[2:]}]}
t2={"title":"振动测试","sample_no":"1#-2#","standard":"参考客户要求","start_date":"2026.05.25",
 "end_date":"2026.05.27","overall_result":"合格","sample_name":"O1D大通DMS一体机",
 "env":"18℃-28℃、25%RH-75%RH","test_date":"2026.05.25-2026.05.27","condition":"随机振动 X/Y/Z","requirement":"功能等级A",
 "equipment":[{"name":"振动试验台","model":"DC-1000-15","mgmt_no":"ME/LAB-015","cal_valid":"2025.11.17-2026.11.16"}],
 "samples":[{"no":"1#","result":"三轴振动前后正常","conclusion":"合格"}],
 "image_groups":[{"title":"X轴图片","images":to_imgs(up2)}]}
proj["tests"]=[t1,t2]
print("save2", post_json("/api/save", proj))

# 3) 生成
res=post_json("/api/generate", proj)
print("generate", res)
