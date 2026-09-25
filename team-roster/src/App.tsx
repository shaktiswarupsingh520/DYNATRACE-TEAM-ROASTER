import {useEffect,useMemo,useState} from 'react';
import {documentsClient} from '@dynatrace-sdk/client-document';
import {Plus,Search,Pencil,Download,X,Lock,RefreshCw,CalendarDays,Users,Trash2,Upload,FileSpreadsheet} from 'lucide-react';
import * as XLSX from 'xlsx';

type Code='G'|'E'|'M'|'W'|'L'|'H';
type Member={id:string;name:string;tpid:string;mobNum:string;role:string;location:string;shift:string;email:string;codes:Record<string,Code>};

const shifts:Record<Code,{label:string;time:string}>={
  G:{label:'General',time:'9:30 AM–7:00 PM'},
  E:{label:'Evening',time:'1:00 PM–10:30 PM'},
  M:{label:'Morning',time:'8:00 AM–5:30 PM'},
  W:{label:'Week Off',time:''},
  L:{label:'Leave',time:''},
  H:{label:'Holiday',time:''}
};

const seed:Member[]=[
  {id:'1',name:'Team Member 1',tpid:'TPID001',mobNum:'',role:'Dynatrace Engineer',location:'Mumbai',shift:'General',email:'',codes:{}},
  {id:'2',name:'Team Member 2',tpid:'TPID002',mobNum:'',role:'Dynatrace Engineer',location:'Mumbai',shift:'General',email:'',codes:{}},
  {id:'3',name:'Team Member 3',tpid:'TPID003',mobNum:'',role:'Dynatrace Engineer',location:'Mumbai',shift:'General',email:'',codes:{}}
];

const getMonthDates=(month:string)=>{
  const [year,monthNumber]=month.split('-').map(Number);
  const count=new Date(year,monthNumber,0).getDate();
  return Array.from({length:count},(_,i)=>{
    const d=new Date(year,monthNumber-1,i+1);
    return {
      key:`${month}-${String(i+1).padStart(2,'0')}`,
      day:d.toLocaleDateString('en-IN',{weekday:'short'}),
      date:`${String(i+1).padStart(2,'0')}-${d.toLocaleDateString('en-IN',{month:'short'})}`
    };
  });
};

const formatMonth=(month:string)=>{
  const [year,monthNumber]=month.split('-').map(Number);
  return new Date(year,monthNumber-1,1).toLocaleDateString('en-IN',{month:'long',year:'numeric'});
};

export default function App(){
  const [members,setMembers]=useState<Member[]>(seed);
  const [meta,setMeta]=useState<any>(null);
  const [loading,setLoading]=useState(true);
  const [error,setError]=useState('');
  const [query,setQuery]=useState('');
  const [open,setOpen]=useState(false);
  const [editing,setEditing]=useState<Member|null>(null);
  const [form,setForm]=useState<Member>(seed[0]);
  const [initialized,setInitialized]=useState(false);
  const [selectedMonth,setSelectedMonth]=useState('2026-09');
  const dates=useMemo(()=>getMonthDates(selectedMonth),[selectedMonth]);

  const canEdit=Boolean(meta?.access?.includes('write'));

  const load=async()=>{
    setLoading(true);
    setError('');
    try{
      const m=await documentsClient.getDocumentMetadata({id:'dynatrace-team-roster'});
      setMeta(m);
      setInitialized(true);

      const r=await documentsClient.downloadDocumentContent({id:m.id});
      const data=await r.get('json') as {members?:Member[]};
      setMembers(Array.isArray(data?.members)?data.members:seed);
    }catch(e:any){
      const status=e?.status ?? e?.response?.status;
      if(status===404){
        setInitialized(false);
        setMeta(null);
      }else{
        setError(e instanceof Error?e.message:'Unable to load the shared roster.');
      }
    }finally{
      setLoading(false);
    }
  };

  useEffect(()=>{void load()},[]);

  const init=async()=>{
    setLoading(true);
    setError('');
    try{
      const content=new Blob(
        [JSON.stringify({members:seed,updatedAt:new Date().toISOString()},null,2)],
        {type:'application/json'}
      );
      const created=await documentsClient.createDocument({
        body:{
          name:'Dynatrace Team Roster',
          type:'dynatrace-team-roster',
          externalId:'dynatrace-team-roster',
          description:'Axis Bank Dynatrace Support Team roster',
          content
        }
      });

      // The create response is authoritative for the newly created document.
      // Reload using the stable external id so access/version metadata is populated.
      if(!created){
        throw new Error('Document creation returned no result.');
      }
      await load();
    }catch(e:any){
      setError(e instanceof Error?e.message:'Initialization failed. Make sure you have document write permission.');
      setLoading(false);
    }
  };

  const importExcel=async(file:File)=>{
    if(!canEdit) return;
    setLoading(true);
    setError('');
    try{
      const buffer=await file.arrayBuffer();
      const workbook=XLSX.read(buffer,{type:'array',cellDates:true});
      const sheetName=workbook.SheetNames.find(n=>n.toLowerCase()==='roster')||workbook.SheetNames[0];
      if(!sheetName) throw new Error('The Excel file does not contain a Roster sheet.');
      const sheet=workbook.Sheets[sheetName];
      const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:true}) as unknown[][];
      const monthCell=rows[2]?.[1];
      const monthRef=sheet.B3;
      let importedMonth='';
      const setMonth=(d:Date)=>{
        if(!Number.isNaN(d.getTime())){
          importedMonth=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
        }
      };
      if(monthCell instanceof Date){
        setMonth(monthCell);
      }else if(monthRef?.t==='n' && typeof monthRef.v==='number'){
        const parsed=XLSX.SSF.parse_date_code(monthRef.v);
        if(parsed) importedMonth=`${parsed.y}-${String(parsed.m).padStart(2,'0')}`;
      }else if(typeof monthCell==='number'){
        const parsed=XLSX.SSF.parse_date_code(monthCell);
        if(parsed) importedMonth=`${parsed.y}-${String(parsed.m).padStart(2,'0')}`;
      }else if(typeof monthCell==='string'){
        const text=monthCell.trim();
        const numeric=Number(text);
        const iso=text.match(/^(\\d{4})-(\\d{1,2})(?:-\\d{1,2})?/);
        const monthYear=text.match(/^(January|February|March|April|May|June|July|August|September|October|November|December)[\\s,]+(\\d{4})$/i);
        if(Number.isFinite(numeric) && numeric>30000 && numeric<60000){
          const parsed=XLSX.SSF.parse_date_code(numeric);
          if(parsed) importedMonth=\`${parsed.y}-${String(parsed.m).padStart(2,'0')}\`;
        }else if(iso){
          importedMonth=\`${iso[1]}-${String(Number(iso[2])).padStart(2,'0')}\`;
        }else if(monthYear){
          const parsed=new Date(\`${monthYear[1]} 1, ${monthYear[2]}\`);
          setMonth(parsed);
        }else{
          const parsed=new Date(text);
          setMonth(parsed);
        }
      }
      if(!/^\\d{4}-(0[1-9]|1[0-2])$/.test(importedMonth)){
        throw new Error('Could not read the Roster Month from cell B3. Please select a month in the Excel template and save it before importing.');
      }
      const importedDates=getMonthDates(importedMonth);
      const header=rows[4]||[];
      if(header[0]!=='Name'||header[1]!=='TPID'||header[2]!=='Mob Num'){
        throw new Error('Invalid Excel format. The first three columns must be Name, TPID and Mob Num.');
      }
      const dataRows=rows.slice(5).filter(row=>String(row?.[0]??'').trim()||String(row?.[1]??'').trim()||String(row?.[2]??'').trim());
      if(dataRows.length===0) throw new Error('No team-member rows were found in the Excel file.');
      if(dataRows.length>50) throw new Error('The Excel file contains more than 50 team members.');
      const existingByKey=new Map(members.map(m=>[
        String(m.tpid||m.name).trim().toLowerCase(),
        m
      ]));
      const imported:Member[]=dataRows.map((row,index)=>{
        const name=String(row[0]??'').trim();
        const tpid=String(row[1]??'').trim();
        const mobNum=String(row[2]??'').trim();
        if(!name) throw new Error(`Row ${index+6}: Name is required.`);
        if(!tpid) throw new Error(`Row ${index+6}: TPID is required for ${name}.`);
        const key=tpid.toLowerCase();
        const previous=existingByKey.get(key)||members.find(m=>m.name.trim().toLowerCase()===name.toLowerCase());
        const codes={...(previous?.codes||{})};
        let firstShift:Code|undefined;
        importedDates.forEach((d,dayIndex)=>{
          const raw=String(row[dayIndex+3]??'').trim().toUpperCase();
          if(raw && !(['G','E','M','W','L','H'] as string[]).includes(raw)){
            throw new Error(`Row ${index+6}, ${d.date}: invalid code "${raw}". Use G, E, M, W, L or H.`);
          }
          if(raw){
            codes[d.key]=raw as Code;
            if(!firstShift && ['G','E','M'].includes(raw)) firstShift=raw as Code;
          }else{
            delete codes[d.key];
          }
        });
        return {
          id:previous?.id||crypto.randomUUID(),
          name,
          tpid,
          mobNum,
          role:previous?.role||'Dynatrace Engineer',
          location:previous?.location||'Mumbai',
          shift:firstShift?shifts[firstShift].label:(previous?.shift||'General'),
          email:previous?.email||'',
          codes
        };
      });
      await writeRoster(imported);
      setSelectedMonth(importedMonth);
      await load();
    }catch(e:any){
      setError(e instanceof Error?e.message:'Excel import failed. Check the template and try again.');
      setLoading(false);
    }
  };

  const writeRoster=async(next:Member[])=>{
    if(!meta?.id || !meta?.version){
      throw new Error('The roster document version is unavailable. Refresh and try again.');
    }
    await documentsClient.updateDocumentContent({
      id:meta.id,
      optimisticLockingVersion:meta.version,
      body:{
        content:new Blob(
          [JSON.stringify({members:next,updatedAt:new Date().toISOString()},null,2)],
          {type:'application/json'}
        )
      }
    });
  };

  const save=async()=>{
    if(!canEdit || !form.name.trim()) return;

    const next=editing
      ? members.map(m=>m.id===editing.id?form:m)
      : [...members,{...form,id:crypto.randomUUID()}];

    try{
      await writeRoster(next);
      setOpen(false);
      await load();
    }catch(e:any){
      setError(e instanceof Error?e.message:'Save failed. Refresh and try again.');
    }
  };

  const remove=async(id:string)=>{
    if(!canEdit || !confirm('Remove this team member?')) return;

    const next=members.filter(m=>m.id!==id);
    try{
      await writeRoster(next);
      await load();
    }catch(e:any){
      setError(e instanceof Error?e.message:'Delete failed. Refresh and try again.');
    }
  };

  const filtered=useMemo(
    ()=>members.filter(m=>Object.values(m).join(' ').toLowerCase().includes(query.toLowerCase())),
    [members,query]
  );

  const exportCsv=()=>{
    const h=['Member','Role','Location','Shift',...dates.map(d=>d.date)];
    const rows=members.map(m=>[
      m.name,m.role,m.location,m.shift,...dates.map(d=>m.codes[d.key]||'')
    ]);
    const csv=[h,...rows].map(r=>r.map(v=>JSON.stringify(v)).join(',')).join('\n');
    const a=document.createElement('a');
    a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));
    a.download='dynatrace-team-roster.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  };

  const openEdit=(m:Member)=>{
    if(!canEdit) return;
    setEditing(m);
    setForm({...m,codes:{...m.codes}});
    setOpen(true);
  };

  const addMember=()=>{
    setEditing(null);
    setForm({
      ...seed[0],
      id:crypto.randomUUID(),
      name:'',
      codes:{}
    });
    setOpen(true);
  };

  return <div className="app">
    <header>
      <div>
        <div className="title">Dynatrace Team Roster</div>
        <div className="subtitle">Support Team • {formatMonth(selectedMonth)}</div>
      </div>

      <div className="header-actions">
        <span className={canEdit?'mode edit':'mode'}>
          {canEdit?<><Pencil size={14}/> Owner edit access</>:<><Lock size={14}/> View only</>}
        </span>
        {canEdit&&<label className="secondary upload-button">
          <Upload size={15}/> Import Excel
          <input
            type="file"
            accept=".xlsx,.xls"
            hidden
            onChange={e=>{
              const file=e.target.files?.[0];
              if(file) void importExcel(file);
              e.currentTarget.value='';
            }}
          />
        </label>}
        <button className="secondary" onClick={exportCsv}>
          <Download size={15}/> Export
        </button>
        {canEdit&&<button onClick={addMember}>
          <Plus size={16}/> Add Member
        </button>}
      </div>
    </header>

    <main>
      {error&&<div className="alert">{error}</div>}

      {loading
        ? <div className="loading"><RefreshCw className="spin"/>Loading shared roster…</div>
        : !initialized
          ? <div className="setup">
              <Users size={32}/>
              <h2>Shared roster not initialized</h2>
              <p>Create the tenant-wide roster once. After initialization, everyone in the tenant can view it and the document owner can edit it.</p>
              <button onClick={init}>Initialize Shared Roster</button>
            </div>
          : <>
              <section className="summary">
                <div><span>Total members</span><b>{members.length}</b></div>
                <div><span>General</span><b>{members.filter(m=>m.shift==='General').length}</b></div>
                <div><span>Morning</span><b>{members.filter(m=>m.shift==='Morning').length}</b></div>
                <div><span>Evening</span><b>{members.filter(m=>m.shift==='Evening').length}</b></div>
                <div><span>View access</span><b>Tenant</b></div>
              </section>

              <section className="toolbar">
                <div className="month-picker">
                  <CalendarDays size={16}/>
                  <label>Month</label>
                  <input type="month" value={selectedMonth} onChange={e=>setSelectedMonth(e.target.value)}/>
                </div>
                <div className="search">
                  <Search size={16}/>
                  <input
                    value={query}
                    onChange={e=>setQuery(e.target.value)}
                    placeholder="Search team member or role…"
                  />
                </div>
                <button className="secondary" onClick={()=>void load()}>
                  <RefreshCw size={15}/> Refresh
                </button>
              </section>

              <section className="roster-card">
                <div className="table-scroll">
                  <table>
                    <thead>
                      <tr>
                        <th className="sticky-name">Team member</th>
                        {dates.map(d=>
                          <th key={d.key}>
                            <b>{d.date}</b>
                            <small>{d.day}</small>
                          </th>
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {filtered.map(m=>
                        <tr key={m.id}>
                          <td className="sticky-name member">
                            <strong>{m.name}</strong>
                            <small>{m.role}</small>
                            <small>{m.tpid} • {m.mobNum||'No mobile'}</small>
                            <small>{m.location} • {m.shift}</small>
                            {canEdit&&
                              <button className="member-edit" title="Edit member" onClick={()=>openEdit(m)}>
                                <Pencil size={12}/>
                              </button>
                            }
                          </td>
                          {dates.map(d=>{
                            const c=m.codes[d.key];
                            return <td
                              key={d.key}
                              className={'code '+(c?`c-${c}`:'empty-code')}
                              onClick={()=>openEdit(m)}
                            >{c||'·'}</td>;
                          })}
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </section>

              <section className="legend">
                <h3><CalendarDays size={16}/> Shift legend</h3>
                {Object.entries(shifts).map(([c,s])=>
                  <div key={c}>
                    <span className={'legend-code c-'+c}>{c}</span>
                    <span>
                      <b>{s.label}</b>{s.time&&<> <small>({s.time})</small></>}
                    </span>
                  </div>
                )}
              </section>
            </>
      }
    </main>

    {open&&canEdit&&
      <div className="overlay">
        <div className="modal">
          <div className="modal-head">
            <div>
              <h2>{editing?'Edit roster':'Add team member'}</h2>
              <small>Owner-only editing</small>
            </div>
            <button className="icon" onClick={()=>setOpen(false)}><X/></button>
          </div>

          <div className="form">
            <label>Name<input value={form.name} onChange={e=>setForm({...form,name:e.target.value})}/></label>
            <label>TPID<input value={form.tpid} onChange={e=>setForm({...form,tpid:e.target.value})}/></label>
            <label>Mob Num<input value={form.mobNum} onChange={e=>setForm({...form,mobNum:e.target.value})}/></label>
            <label>Role<input value={form.role} onChange={e=>setForm({...form,role:e.target.value})}/></label>
            <label>Location<input value={form.location} onChange={e=>setForm({...form,location:e.target.value})}/></label>
            <label>Default shift
              <select value={form.shift} onChange={e=>setForm({...form,shift:e.target.value})}>
                <option>General</option>
                <option>Morning</option>
                <option>Evening</option>
              </select>
            </label>
            <label>Email<input value={form.email} onChange={e=>setForm({...form,email:e.target.value})}/></label>

            <div className="days">
              <b>Daily roster codes</b>
              <div className="day-grid">
                {dates.map(d=>
                  <label key={d.key}>
                    <span>{d.date}</span>
                    <select
                      value={form.codes[d.key]||''}
                      onChange={e=>setForm({
                        ...form,
                        codes:{...form.codes,[d.key]:e.target.value as Code}
                      })}
                    >
                      <option value="">—</option>
                      {Object.keys(shifts).map(c=><option key={c}>{c}</option>)}
                    </select>
                  </label>
                )}
              </div>
            </div>
          </div>

          <div className="modal-foot">
            {editing&&<button
              className="danger"
              onClick={()=>void remove(editing.id)}
            >
              <Trash2 size={14}/> Remove
            </button>}
            <button className="secondary" onClick={()=>setOpen(false)}>Cancel</button>
            <button onClick={()=>void save()}>{editing?'Save Changes':'Add Member'}</button>
          </div>
        </div>
      </div>
    }
  </div>
}
