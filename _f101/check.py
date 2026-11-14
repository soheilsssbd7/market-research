import importlib.util, os, glob, re
BASE=os.path.dirname(os.path.abspath(__file__))
tot=0
for f in sorted(glob.glob(os.path.join(BASE,'p*.py'))):
    spec=importlib.util.spec_from_file_location('m',f); m=importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
    w=0; h1=0; h2=0; t=0; qc=0
    for raw in m.P:
        k,_,txt=raw.partition('|')
        if k=='@H1': h1+=1
        if k=='@H2': h2+=1
        if k=='@T': t+=1
        if k=='@QC': qc+=1
        if k not in ('@PAGEBREAK',): w+=len(txt.split())
    tot+=w
    print(os.path.basename(f), w)
print('TOTAL', tot)
