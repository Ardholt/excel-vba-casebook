"""Score a completed extraction against a separate key; cannot fill output values."""
import csv,json,sys
from pathlib import Path

def evaluate(output, key):
    output=Path(output)
    with (output/'fields.csv').open() as f: rows=list(csv.DictReader(f))
    with Path(key).open() as f: truth=list(csv.DictReader(f))
    expected={(r['image'],r['field']):r['value'] for r in truth}
    actual={(r['image'],r['field']) for r in rows}
    if len(expected)!=len(truth) or len(actual)!=len(rows) or set(expected)!=actual:
        raise ValueError('Answer key must cover each image and field exactly once')
    report={'denominator':len(rows),'scope':'Only the supplied image set, compared with the separate answer key.'}
    details=[]
    for phase,col in [('raw','raw_value'),('reviewed','value')]:
        count={'correct':0,'unresolved':0,'errors':0}
        for r in rows:
            answer=expected[r['image'],r['field']]
            status='unresolved' if not r[col] else 'correct' if r[col]==answer else 'errors'
            count[status]+=1
            details.append({'phase':phase,'image':r['image'],'field':r['field'],'expected':answer,'actual':r[col],'result':status})
        report[phase]=count
    (output/'accuracy.json').write_text(json.dumps(report,indent=2)+'\n')
    with (output/'accuracy-fields.csv').open('w',newline='') as f:
        w=csv.DictWriter(f,fieldnames=list(details[0]));w.writeheader();w.writerows(details)
    return report
if __name__=='__main__': evaluate(sys.argv[1],sys.argv[2])
