#!/usr/bin/env python3
"""UNO worker; run with a Python that can import the installed LibreOffice uno."""
import argparse
import json
import subprocess
import tempfile
import time
import uuid
from pathlib import Path


def main():
    ap=argparse.ArgumentParser();ap.add_argument('source');ap.add_argument('output');ap.add_argument('report');ap.add_argument('--soffice',default='libreoffice');a=ap.parse_args()
    result={'status':'failed','engine':'LibreOffice UNO','fields_updated':False,'indexes_updated':False,'errors':[]}
    proc=None;doc=None;desktop=None
    try:
        import uno
        from com.sun.star.beans import PropertyValue
        def prop(name,val):
            p=PropertyValue();p.Name=name;p.Value=val;return p
        with tempfile.TemporaryDirectory(prefix='dlr-lo-') as tmp:
            pipe='dlr_'+uuid.uuid4().hex
            proc=subprocess.Popen([a.soffice,'-env:UserInstallation='+Path(tmp).as_uri(),'--headless','--nologo','--nodefault','--nofirststartwizard','--accept=pipe,name='+pipe+';urp;StarOffice.ServiceManager'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
            local=uno.getComponentContext();resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
            ctx=None
            for _ in range(100):
                try:ctx=resolver.resolve('uno:pipe,name='+pipe+';urp;StarOffice.ComponentContext');break
                except Exception:time.sleep(.2)
            if ctx is None:raise RuntimeError('Cannot connect to LibreOffice UNO')
            desktop=ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
            doc=desktop.loadComponentFromURL(Path(a.source).resolve().as_uri(),'_blank',0,(prop('Hidden',True),prop('ReadOnly',False),prop('UpdateDocMode',uno.getConstantByName('com.sun.star.document.UpdateDocMode.NO_UPDATE')),prop('MacroExecutionMode',uno.getConstantByName('com.sun.star.document.MacroExecMode.NEVER_EXECUTE'))))
            if doc is None:raise RuntimeError('LibreOffice failed to load DOCX')
            for _ in range(2):
                doc.getTextFields().refresh()
                indexes=doc.getDocumentIndexes()
                for i in range(indexes.getCount()):indexes.getByIndex(i).update()
                doc.refresh()
            doc.storeAsURL(Path(a.output).resolve().as_uri(),(prop('FilterName','Office Open XML Text'),prop('Overwrite',False)))
            doc.close(True);doc=None
            result.update(status='passed',fields_updated=True,indexes_updated=True)
    except Exception as exc:result['errors']=[str(exc)]
    finally:
        if doc:
            try:doc.close(True)
            except Exception:pass
        if desktop:
            try:desktop.terminate()
            except Exception:pass
        if proc:
            try:proc.wait(timeout=10)
            except subprocess.TimeoutExpired:proc.terminate()
        Path(a.report).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    return 0 if result['status']=='passed' else 2


if __name__=='__main__':raise SystemExit(main())
