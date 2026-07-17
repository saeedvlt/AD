import pandas as pd
from openpyxl import load_workbook

MONTH_MAP={
    "Jan":"01 - Jan","Feb":"02 - Feb","Mar":"03 - Mar","Apr":"04 - Apr",
    "May":"05 - May","Jun":"06 - Jun","Jul":"07 - Jul","Aug":"08 - Aug",
    "Sep":"09 - Sep","Sept":"09 - Sep","Oct":"10 - Oct","Nov":"11 - Nov","Dec":"12 - Dec"
}

IGNORE_SHEETS={"Instructions","Totals","Total PP Indirect","By GL Account","Sheet1"}

FUNCTION_NAMES={
    "Operations","HR","Finance","Sales","Engineering","Purchasing",
    "Quality","IT","Management","Processing Planning","Maintenance","Production"
}

def find_header_row(ws):
    months=set(MONTH_MAP.keys())
    for r in range(1,min(ws.max_row,40)+1):
        values=[("" if ws.cell(r,c).value is None else str(ws.cell(r,c).value).strip()) for c in range(1,ws.max_column+1)]
        if sum(v in months for v in values)>=6:
            return r
    return None

def month_columns(ws,header):
    cols={}
    for c in range(1,ws.max_column+1):
        v=ws.cell(header,c).value
        if v is None:
            continue
        v=str(v).strip()
        if v in MONTH_MAP:
            cols[MONTH_MAP[v]]=c
    return cols

def convert(uploaded_file):
    wb=load_workbook(uploaded_file,data_only=True)
    records=[]
    for sheet in [s for s in wb.sheetnames if s not in IGNORE_SHEETS]:
        ws=wb[sheet]
        header=find_header_row(ws)
        if header is None:
            continue
        months=month_columns(ws,header)
        current_gl=None
        current_function=None
        for r in range(header+1,ws.max_row+1):
            a=ws.cell(r,1).value
            b=ws.cell(r,2).value
            c=ws.cell(r,3).value
            text="" if a is None else str(a).strip()
            if " - " in text:
                current_gl=text
                current_function=None
                continue
            if text in FUNCTION_NAMES:
                current_function=text
                continue
            values={}
            has_data=False
            for month,col in months.items():
                val=ws.cell(r,col).value
                if val is None:
                    val=0
                values[month]=val
                if isinstance(val,(int,float)) and val!=0:
                    has_data=True
            if not has_data:
                continue
            item=text if text else b
            for month,val in values.items():
                records.append({
                    "Expense Category":sheet,
                    "GL Account":current_gl,
                    "Functional Unit":current_function,
                    "Item":item,
                    "Description":c,
                    "Month":month,
                    "Amount":val
                })
    return pd.DataFrame(records)
