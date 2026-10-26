# -*- coding: utf-8 -*-
import io,json,tempfile,webbrowser,threading,re
from pathlib import Path
from http.server import ThreadingHTTPServer,BaseHTTPRequestHandler
from email.parser import BytesParser
from email.policy import default
from openpyxl import load_workbook
from compare_multi import run_multi
ROOT=Path(__file__).parent
LATEST={'excel':None}
def records(ws):
    it=ws.values;headers=next(it)
    return [dict(zip(headers,row)) for row in it]
class Handler(BaseHTTPRequestHandler):
    def send(self,data,typ='application/json',status=200,filename=None):
        self.send_response(status);self.send_header('Content-Type',typ);self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store, no-cache, must-revalidate');self.send_header('Pragma','no-cache')
        if filename:self.send_header('Content-Disposition',f'attachment; filename="{filename}"')
        self.end_headers();self.wfile.write(data)
    def do_GET(self):
        if self.path=='/':self.send((ROOT/'dashboard.html').read_bytes(),'text/html; charset=utf-8')
        elif self.path=='/download' and LATEST['excel']:
            self.send(LATEST['excel'],'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',filename='PO_Multi_Season_Result.xlsx')
        else:self.send(b'Not found','text/plain',404)
    def do_POST(self):
        if self.path!='/compare':return self.send(b'Not found','text/plain',404)
        try:
            size=int(self.headers.get('Content-Length','0'))
            if size<=0 or size>110_000_000:raise ValueError('전체 업로드는 110MB 이하여야 합니다.')
            body=self.rfile.read(size)
            msg=BytesParser(policy=default).parsebytes(b'Content-Type: '+self.headers['Content-Type'].encode()+b'\r\nMIME-Version: 1.0\r\n\r\n'+body)
            parts={'order':[],'po':[]}
            for part in msg.iter_parts():
                field=part.get_param('name',header='content-disposition')
                if field in parts:
                    content=part.get_payload(decode=True)
                    if len(content)>25_000_000:raise ValueError('개별 파일은 25MB 이하여야 합니다.')
                    filename=part.get_filename() or 'uploaded.xlsx'
                    if not filename.lower().endswith('.xlsx'):raise ValueError('.xlsx 파일만 업로드할 수 있습니다.')
                    parts[field].append((filename,content))
            if not parts['order'] or not parts['po']:raise ValueError('각 유형별로 파일을 1개 이상 선택하세요.')
            with tempfile.TemporaryDirectory() as d:
                paths={}
                for field,files in parts.items():
                    paths[field]=[]
                    for idx,(filename,content) in enumerate(files):
                        safe=re.sub(r'[^\w. -]','_',filename)
                        p=Path(d)/(f'{field}_{idx:03d}_{safe}');p.write_bytes(content);paths[field].append(p)
                out=Path(d)/'result.xlsx'
                run_multi(paths['order'],paths['po'],out)
                excel=out.read_bytes()
                wb=load_workbook(io.BytesIO(excel),read_only=True,data_only=True)
                sheets=['Dashboard','Discrepancy','Factory Summary','Season Summary','Division Summary','Style Summary','Month Summary','Upload Audit','Color Review','Excluded Orders','Order Source','PO Source']
                data={k:records(wb[k]) for k in sheets};wb.close()
                # restore original names in audit
                for row in data['Upload Audit']:
                    f=row['File']
                    row['File']=re.sub(r'^(ORDER: |PO: )(?:order|po)_\d{3}_',r'\1',f)
            LATEST['excel']=excel
            self.send(json.dumps(data,ensure_ascii=False,default=str).encode('utf-8'))
        except Exception as e:self.send(json.dumps({'error':str(e)},ensure_ascii=False).encode('utf-8'),'application/json',400)
    def log_message(self,*args):pass
if __name__=='__main__':
    import argparse
    p=argparse.ArgumentParser();p.add_argument('--port',type=int,default=8880);args=p.parse_args()
    server=ThreadingHTTPServer(('127.0.0.1',args.port),Handler)
    url=f'http://127.0.0.1:{args.port}'
    print('Multi-Season PO Dashboard:',url,flush=True)
    threading.Timer(0.8,lambda:webbrowser.open(url)).start()
    try:server.serve_forever()
    except KeyboardInterrupt:pass
