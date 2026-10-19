// Сборка красиво оформленной книги Excel. Чистая функция: XLSX (xlsx-js-style) + данные -> workbook
function buildBook(X, d){
  const BLUE='262626', GRAY='404040', LIGHT='F2F2F2', ZEBRA='FFFFFF', LINE='D9D9D9', RED='262626', GREEN='595959';
  const border={top:{style:'thin',color:{rgb:LINE}},bottom:{style:'thin',color:{rgb:LINE}},left:{style:'thin',color:{rgb:LINE}},right:{style:'thin',color:{rgb:LINE}}};
  const font=(o)=>Object.assign({name:'Arial',sz:11,color:{rgb:'1D1F24'}},o||{});
  const S={
    title:{font:font({sz:15,bold:true,color:{rgb:BLUE}}),alignment:{vertical:'center'}},
    sub:{font:font({sz:10,color:{rgb:'62666F'}})},
    head:{font:font({bold:true,color:{rgb:'262626'}}),fill:{fgColor:{rgb:'E7E6E6'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border},
    label:{font:font({bold:true,color:{rgb:GRAY}}),fill:{fgColor:{rgb:'FFFFFF'}},border,alignment:{vertical:'center',indent:1}},
    text:{font:font(),border,alignment:{vertical:'center',wrapText:true,indent:1}},
    num:{font:font(),border,numFmt:'#,##0',alignment:{horizontal:'right',vertical:'center'}},
    date:{font:font(),border,numFmt:'yyyy.mm.dd',alignment:{horizontal:'center',vertical:'center'}},
    total:{font:font({bold:true,color:{rgb:'262626'}}),fill:{fgColor:{rgb:LIGHT}},border:Object.assign({},border,{top:{style:'medium',color:{rgb:'7F7F7F'}}}),numFmt:'#,##0',alignment:{horizontal:'right',vertical:'center'}},
    totalL:{font:font({bold:true,color:{rgb:'262626'}}),fill:{fgColor:{rgb:LIGHT}},border:Object.assign({},border,{top:{style:'medium',color:{rgb:'7F7F7F'}}}),alignment:{vertical:'center',indent:1}},
  };
  const zebra=(st,i)=>i%2?Object.assign({},st,{fill:{fgColor:{rgb:ZEBRA}}}):st;
  const cell=(v,s,t)=>({v,t:t||(typeof v==='number'?'n':'s'),s});
  const put=(ws,r,c,o)=>{ws[X.utils.encode_cell({r,c})]=o};
  const finish=(ws,rows,cols,head)=>{ws['!ref']=X.utils.encode_range({s:{r:0,c:0},e:{r:rows-1,c:cols-1}});ws['!rows']=Array.from({length:rows},(_,i)=>({hpt:i===head?30:21}))};
  const wb=X.utils.book_new();

  // ---- Лист 1: Сводка ----
  {
    const ws={}, n=d.staff.length; let r=0;
    put(ws,r,0,cell('Расходы команды — сводка',S.title)); r++;
    put(ws,r,0,cell(`Весь период, данные на ${d.today}`,S.sub)); r+=2;
    const hr=r;
    put(ws,r,0,cell('',S.head));
    d.staff.forEach((s,i)=>put(ws,r,1+i,cell(s.name,S.head)));
    put(ws,r,1+n,cell('Итого',S.head)); r++;
    const rowsDef=[
      ['Бюджет',d.staff.map(s=>s.budget||null)],
      ['Потрачено',d.staff.map(s=>s.spent)],
      ['Выдано',d.staff.map(s=>s.paid)],
      ['Остаток',d.staff.map(s=>s.budget?s.budget-s.due:-s.due)],
    ];
    rowsDef.forEach(([lab,vals],k)=>{
      put(ws,r,0,cell(lab,zebra(S.label,k)));
      vals.forEach((v,i)=>{
        if(v===null){put(ws,r,1+i,cell('–',zebra(Object.assign({},S.text,{alignment:{horizontal:'center'}}),k)));return}
        const st=Object.assign({},zebra(S.num,k));
        if(lab==='Остаток'&&v<0)st.font=font({bold:true,color:{rgb:RED}});
        if(lab==='Выдано'&&v>0)st.font=font({color:{rgb:GREEN}});
        put(ws,r,1+i,cell(v,st));
      });
      const sum=vals.reduce((a,v)=>a+(v||0),0);
      put(ws,r,1+n,cell(sum,zebra(Object.assign({},S.num,{font:font({bold:true})}),k)));
      r++;
    });
    put(ws,r,0,cell('К выдаче',S.totalL));
    d.staff.forEach((s,i)=>put(ws,r,1+i,cell(s.due,S.total)));
    put(ws,r,1+n,cell(d.staff.reduce((a,s)=>a+s.due,0),S.total)); r++;
    finish(ws,r,n+2,hr);
    ws['!cols']=[{wch:14},...d.staff.map(()=>({wch:15})),{wch:15}];
    ws['!rows'][0]={hpt:26};ws['!rows'][2]={hpt:8};
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:n+1}},{s:{r:1,c:0},e:{r:1,c:n+1}}];
    X.utils.book_append_sheet(wb,ws,'Сводка');
  }

  // ---- Лист: По неделям (пн–вс), невыданное переносится на следующую неделю ----
  if(d.weeks&&d.weeks.length){
    const ws={}; let r=0;
    const H=['Неделя','Сотрудник','Бюджет','Не выдано с прошлых недель','Потрачено за неделю','Итого к выдаче','Выдано','Дата выдачи','Не выдано на конец недели','Статус'];
    put(ws,r,0,cell('Расходы и выдачи по неделям',S.title)); r++;
    put(ws,r,0,cell('Неделя — с понедельника по воскресенье. Невыданная сумма не обнуляется: она остаётся долгом и прибавляется к следующей неделе.',S.sub)); r+=2;
    const hr=r; H.forEach((h,c)=>put(ws,r,c,cell(h,S.head))); r++;
    const wk={font:font({bold:true,sz:12,color:{rgb:'262626'}}),fill:{fgColor:{rgb:'D9D9D9'}},border,alignment:{vertical:'center',indent:1}};
    const stc={'Выдано':'8A8A8A','Частично':'3A3A3A','Не выдано':'111111','Неделя идёт':'6B6B6B'};
    const heights={}, wkMerges=[];
    d.weeks.forEach(w=>{
      heights[r]=22; wkMerges.push({s:{r,c:0},e:{r,c:H.length-1}}); for(let c=0;c<H.length;c++)put(ws,r,c,cell(c===0?w.label+(w.current?' · текущая':''):'',wk)); r++;
      w.rows.forEach((x,k)=>{
        put(ws,r,0,cell('',zebra(S.text,k)));
        put(ws,r,1,cell(x.name,zebra(S.text,k)));
        put(ws,r,2,x.budget?cell(x.budget,zebra(S.num,k)):cell('–',zebra(Object.assign({},S.text,{alignment:{horizontal:'center'}}),k)));
        const carrySt=Object.assign({},zebra(S.num,k)); if(x.carry>0)carrySt.font=font({color:{rgb:RED}});
        put(ws,r,3,cell(x.carry,carrySt));
        put(ws,r,4,cell(x.spent,zebra(S.num,k)));
        put(ws,r,5,cell(x.total,Object.assign({},zebra(S.num,k),{font:font({bold:true})})));
        const pSt=Object.assign({},zebra(S.num,k)); if(x.paid>0)pSt.font=font({color:{rgb:GREEN}});
        put(ws,r,6,cell(x.paid,pSt));
        put(ws,r,7,cell(x.paidDates,zebra(Object.assign({},S.text,{alignment:{horizontal:'center',vertical:'center'}}),k)));
        const rSt=Object.assign({},zebra(S.num,k)); if(x.rest>0)rSt.font=font({bold:true,color:{rgb:RED}});
        put(ws,r,8,cell(x.rest,rSt));
        put(ws,r,9,cell(x.status,Object.assign({},zebra(S.text,k),{font:font({bold:true,color:{rgb:stc[x.status]||'1D1F24'}})})));
        r++;
      });
      const sum=f=>w.rows.reduce((a,x)=>a+x[f],0);
      put(ws,r,0,cell('',S.totalL));put(ws,r,1,cell('Итого за неделю',S.totalL));put(ws,r,2,cell('',S.totalL));
      put(ws,r,3,cell(sum('carry'),S.total));put(ws,r,4,cell(sum('spent'),S.total));put(ws,r,5,cell(sum('total'),S.total));
      put(ws,r,6,cell(sum('paid'),S.total));put(ws,r,7,cell('',S.totalL));put(ws,r,8,cell(sum('rest'),S.total));put(ws,r,9,cell('',S.totalL));
      r++; heights[r]=10; r++;
    });
    finish(ws,r,H.length,hr);
    Object.entries(heights).forEach(([i,h])=>{ws['!rows'][i]={hpt:h}});
    ws['!rows'][0]={hpt:26};ws['!rows'][2]={hpt:8};ws['!rows'][hr]={hpt:44};
    ws['!cols']=[{wch:17},{wch:16},{wch:11},{wch:14},{wch:13},{wch:13},{wch:12},{wch:12},{wch:14},{wch:13}];
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:9}},{s:{r:1,c:0},e:{r:1,c:9}},...wkMerges];
    X.utils.book_append_sheet(wb,ws,'По неделям');
  }

  // ---- Лист 2: Расходы ----
  {
    const ws={}; let r=0;
    const H=['Дата','Сотрудник','Категория','Сумма, ₩','Комментарий','Не возмещается'];
    H.forEach((h,c)=>put(ws,r,c,cell(h,S.head))); r++;
    const dim=st=>e=>e.rej?Object.assign({},st,{font:font({color:{rgb:'8C8C8C'},strike:st===S.num})}):st;
    d.expenses.forEach((e,k)=>{
      const g=st=>dim(zebra(st,k))(e);
      put(ws,r,0,cell(e.date,g(S.date),'n'));
      put(ws,r,1,cell(e.name,g(S.text)));
      put(ws,r,2,cell(e.cat,g(S.text)));
      put(ws,r,3,cell(e.amount,e.rej?Object.assign({},zebra(S.num,k),{font:font({color:{rgb:'8C8C8C'},strike:true})}):zebra(S.num,k)));
      put(ws,r,4,cell(e.note||'',g(S.text)));
      put(ws,r,5,cell(e.rej||'',zebra(S.text,k)));
      r++;
    });
    const ok=d.expenses.filter(e=>!e.rej);
    put(ws,r,0,cell('Итого к возмещению',S.totalL));
    put(ws,r,1,cell(`записей: ${ok.length}`,S.totalL));
    put(ws,r,2,cell('',S.totalL));
    put(ws,r,3,cell(ok.reduce((a,e)=>a+e.amount,0),S.total));
    put(ws,r,4,cell('',S.totalL)); put(ws,r,5,cell('',S.totalL)); r++;
    finish(ws,r,6,0);
    ws['!cols']=[{wch:12},{wch:18},{wch:14},{wch:14},{wch:44},{wch:20}];
    if(d.expenses.length)ws['!autofilter']={ref:X.utils.encode_range({s:{r:0,c:0},e:{r:d.expenses.length,c:5}})};
    X.utils.book_append_sheet(wb,ws,'Расходы');
  }

  // ---- Лист 3: Выдачи ----
  {
    const ws={}; let r=0;
    ['Дата','Сотрудник','Выдано, ₩'].forEach((h,c)=>put(ws,r,c,cell(h,S.head))); r++;
    d.payments.forEach((p,k)=>{
      put(ws,r,0,cell(p.date,zebra(S.date,k),'n'));
      put(ws,r,1,cell(p.name,zebra(S.text,k)));
      put(ws,r,2,cell(p.amount,zebra(S.num,k)));
      r++;
    });
    put(ws,r,0,cell('Итого',S.totalL));put(ws,r,1,cell('',S.totalL));
    put(ws,r,2,cell(d.payments.reduce((a,p)=>a+p.amount,0),S.total)); r++;
    finish(ws,r,3,0);
    ws['!cols']=[{wch:12},{wch:18},{wch:14}];
    X.utils.book_append_sheet(wb,ws,'Выдачи');
  }
  return wb;
}

// Недели пн–вс по дате расхода/выдачи; невыданное переносится вперёд
function monOf(iso){const [y,m,d]=iso.split('-').map(Number);const dt=new Date(y,m-1,d);dt.setDate(dt.getDate()-((dt.getDay()+6)%7));return dt}
function weeklyLedger(){
  const items=[...expenses.map(e=>dateOf(e)),...payments.map(p=>p.date).filter(Boolean)];
  if(!items.length)return [];
  const first=monOf(items.sort()[0]), cur=monOf(isoDay(0)), weeks=[];
  const carry={}; staff.forEach(s=>carry[s.id]=0);
  for(let w=new Date(first);w<=cur;w.setDate(w.getDate()+7)){
    const from=isoOf(w.getTime()), end=new Date(w);end.setDate(end.getDate()+6);const to=isoOf(end.getTime());
    const isCur=from===isoOf(cur.getTime()), rows=[];
    for(const s of staff){
      const spent=expenses.filter(e=>e.staff===s.id&&!e.rej&&dateOf(e)>=from&&dateOf(e)<=to).reduce((a,e)=>a+e.amount,0);
      const ps=payments.filter(p=>p.staff===s.id&&p.date>=from&&p.date<=to);
      const paid=ps.reduce((a,p)=>a+p.amount,0);
      const c=carry[s.id]||0, total=c+spent, rest=total-paid;
      carry[s.id]=rest;
      if(!c&&!spent&&!paid)continue;
      const status=rest<=0?'Выдано':isCur?'Неделя идёт':paid>0?'Частично':'Не выдано';
      rows.push({name:s.name,budget:s.budget,carry:c,spent,total,paid,paidDates:[...new Set(ps.map(p=>dmy(p.date)))].join(', '),rest,status});
    }
    if(rows.length)weeks.push({label:`${ymd(w.getTime())} – ${ymd(end.getTime())}`,current:isCur,rows});
  }
  return weeks.reverse(); // новые недели сверху
}

// Отчёт по форме «현금 지출 현황 - ГГГГ년» (как в шаблоне налоговой). 잔액 оставляем пустым.
function buildTaxBook(X, year, rows){
  const thin={style:'thin',color:{rgb:'000000'}}, dbl={style:'double',color:{rgb:'000000'}};
  const B=(left)=>({top:thin,bottom:thin,right:thin,left:left||thin});
  const f=(o)=>Object.assign({name:'맑은 고딕',sz:10,color:{rgb:'000000'}},o||{});
  const num='_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
  const S={
    title:{font:f({sz:14,bold:true}),alignment:{horizontal:'center',vertical:'center'}},
    warn:{font:f({bold:true,sz:9}),fill:{fgColor:{rgb:'FFFF00'}},alignment:{horizontal:'center',vertical:'center',wrapText:true}},
    head:{font:f({bold:true}),fill:{fgColor:{rgb:'D9D9D9'}},alignment:{horizontal:'center',vertical:'center'},border:B()},
    date:{font:f(),numFmt:'mm"월" dd"일"',alignment:{horizontal:'center',vertical:'center'},border:B(dbl)},
    text:{font:f(),alignment:{horizontal:'left',vertical:'center',wrapText:true},border:B()},
    num:{font:f(),numFmt:num,alignment:{horizontal:'center',vertical:'center'},border:B()},
    memo:{font:f(),alignment:{vertical:'center',wrapText:true},border:B()},
    tot:{font:f({bold:true}),fill:{fgColor:{rgb:'F2F2F2'}},numFmt:num,alignment:{horizontal:'center',vertical:'center'},border:B()},
    totL:{font:f({bold:true}),fill:{fgColor:{rgb:'F2F2F2'}},alignment:{horizontal:'center',vertical:'center'},border:B(dbl)},
  };
  const ws={}, put=(r,c,v,s,t)=>{ws[X.utils.encode_cell({r,c})]={v,t:t||(typeof v==='number'?'n':'s'),s}};
  put(0,0,`현금 지출 현황 - ${year}년 `,S.title); for(let c=1;c<6;c++)put(0,c,'',{});
  put(1,0,'영수증등 증빙은 필히 보관하시고 3만원 초과분은 현금영수증, 개인카드사용만 가산세 면제(법인카드,법인사업자현금영수증은 표기 하지 마세요)',S.warn); for(let c=1;c<6;c++)put(1,c,'',S.warn);
  ['일자','상 호 외','내역','지출','잔액','비고'].forEach((h,c)=>put(2,c,h,S.head));
  let r=3;
  rows.forEach(x=>{
    put(r,0,x.date,S.date,'n'); put(r,1,x.place,S.text); put(r,2,x.detail,S.text);
    put(r,3,x.amount,S.num); put(r,4,'',S.num); put(r,5,x.memo,S.memo); r++;
  });
  const minRows=Math.max(r,56);                // пустые строки с рамками, как в шаблоне
  for(;r<minRows;r++){ put(r,0,'',S.date); put(r,1,'',S.text); put(r,2,'',S.text); put(r,3,'',S.num); put(r,4,'',S.num); put(r,5,'',S.memo); }
  put(r,0,'합계',S.totL); put(r,1,'',S.tot); put(r,2,'',S.tot);
  put(r,3,rows.reduce((a,x)=>a+x.amount,0),S.tot); put(r,4,'',S.tot); put(r,5,'',S.tot); r++;
  ws['!ref']=X.utils.encode_range({s:{r:0,c:0},e:{r:r-1,c:5}});
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:5}},{s:{r:1,c:0},e:{r:1,c:5}}];
  ws['!cols']=[{wch:9.6},{wch:16.5},{wch:33.7},{wch:12.2},{wch:14.5},{wch:29.9}];
  ws['!rows']=[{hpt:24},{hpt:30},{hpt:20}];
  const wb=X.utils.book_new(); X.utils.book_append_sheet(wb,ws,`현금-${year}년`); return wb;
}
