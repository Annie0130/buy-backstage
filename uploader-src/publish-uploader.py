# 把 feishu-uploader 的程式檔放到 buy-backstage/tools/feishu-uploader（Netlify 會發佈），產生 manifest.json
import os, sys, json, hashlib, shutil, datetime
HERE=os.path.dirname(os.path.abspath(__file__))
# 原始碼在 buy-backstage/uploader-src；發佈到 buy-backstage/tools/feishu-uploader（使用者的「更新程式」會從這裡下載）
SRC=HERE; DST=os.path.join(HERE,'..','tools','feishu-uploader')
note=sys.argv[1] if len(sys.argv)>1 else ''
SKIP_DIRS={'node_modules','商品照片','chrome-profile','給你','舊版備份'}
SKIP_FILES={'publish-uploader.py','交接說明.md','★控制台.bat','★一鍵上架.bat','★控制台.command','更新程式.bat','更新程式.command','補照片清單.txt','版本.txt','airtable-token.txt','飛鼠新增商品網址.txt','平台分類選項.txt','package-lock.json'}
files=[]
for root,dirs,fs in os.walk(SRC):
    dirs[:]=[d for d in dirs if d not in SKIP_DIRS]
    for f in sorted(fs):
        if f in SKIP_FILES or f.lower().endswith(('.csv','.png','.zip')) or f=='照片分組預覽.html' or f.startswith('.'): continue
        files.append(os.path.relpath(os.path.join(root,f),SRC))
shutil.rmtree(DST,ignore_errors=True); os.makedirs(DST)
out=[]
for i,rel in enumerate(sorted(files)):
    b=open(os.path.join(SRC,rel),'rb').read()
    ext=os.path.splitext(rel)[1].lower() or '.bin'
    name=f'f{i:02d}{ext}.txt' if ext in ('.bat','.js','.json','.txt','.html','.command') else f'f{i:02d}{ext}'
    open(os.path.join(DST,name),'wb').write(b)
    out.append({'path':rel.replace(os.sep,'/'),'url':name,'sha256':hashlib.sha256(b).hexdigest()})
shutil.copy(os.path.join(SRC,'update.js'),os.path.join(DST,'update.js.txt'))  # 固定網址，給第一次安裝用
version=datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime('%Y%m%d-%H%M')
remove=['★一鍵上架（照片＋建立＋蝦皮）.bat','3_待辦清單.bat','0_拍照順序.bat','0_照片分組.bat','1_飛鼠上架（存草稿）.bat','2_蝦皮補資料.bat','4_草稿補照片並發佈.bat','5_單獨補照片.bat','其他功能/5_單獨補照片.bat','圖片欄位診斷.txt','欄位診斷.txt']
json.dump({'version':version,'note':note,'files':out,'remove':remove},open(os.path.join(DST,'manifest.json'),'w',encoding='utf-8'),ensure_ascii=False,indent=1)
print(version,len(out),'files'); [print(' ',o['path'],'→',o['url']) for o in out]
